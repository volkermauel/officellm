using System.Text.Json;
using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tools;

/// <summary>
/// Handles MCP tool definition generation and command dispatch.
/// Uses instance-aware routing with a fixed tool list.
/// </summary>
public static class McpToolEngine
{
    private static readonly InstanceRegistry _registry = new();
    private static readonly CommandStore _commandStore = new();

    /// <summary>
    /// Gets all MCP tool definitions (fixed list, instance selected via parameter).
    /// </summary>
    public static object[] GetToolDefinitions() => [
        new
        {
            name = "office_get_active_app",
            description = "Return active Office host, document name and selection metadata.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>(),
                required = Array.Empty<string>()
            }
        },
        new
        {
            name = "powerpoint_get_deck_outline",
            description = "Returns slide titles, text placeholders, notes metadata and slide order.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "Target instance ID (e.g. 'powerpoint_1'). Omit to use the most recently registered instance." },
                    ["includeSpeakerNotes"] = new { type = "boolean", description = "Include speaker notes in the outline", default_value = false },
                    ["includeHiddenSlides"] = new { type = "boolean", description = "Include hidden slides in the outline", default_value = false }
                },
                required = Array.Empty<string>()
            }
        },
        new
        {
            name = "powerpoint_get_slide",
            description = "Return text, notes and shape metadata for one slide.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "Target instance ID (e.g. 'powerpoint_1'). Omit to use the most recently registered instance." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" }
                },
                required = new[] { "slideIndex" }
            }
        },
        new
        {
            name = "powerpoint_update_shape_text",
            description = "Update a specific shape's text with preview. Requires user confirmation before applying.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "Target instance ID (e.g. 'powerpoint_1'). Omit to use the most recently registered instance." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape name/ID on the slide" },
                    ["text"] = new { type = "string", description = "New text content for the shape" }
                },
                required = new[] { "slideIndex", "shapeId", "text" }
            }
        },
        new
        {
            name = "powerpoint_update_speaker_notes",
            description = "Create or update speaker notes for selected slides.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "Target instance ID (e.g. 'powerpoint_1'). Omit to use the most recently registered instance." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["notes"] = new { type = "string", description = "Speaker notes text" }
                },
                required = new[] { "slideIndex", "notes" }
            }
        }
    ];

    /// <summary>
    /// Executes an MCP tool call. Routes to the appropriate add-in instance.
    /// </summary>
    public static async Task<object> ExecuteTool(string name, JsonElement? args)
    {
        // Extract instanceId from args, default to most recent registered instance
        string? instanceId = null;
        if (args.HasValue && args.Value.TryGetProperty("instanceId", out var iid))
            instanceId = iid.GetString();

        // If no instanceId specified, use the most recent active instance
        if (string.IsNullOrEmpty(instanceId))
        {
            var instances = _registry.GetActiveInstances();
            if (!instances.Any())
            {
                return new
                {
                    content = new[] { new { type = "text", text = "No Office instances registered. Open PowerPoint and load the add-in." } },
                    isError = true
                };
            }
            instanceId = instances.Last().InstanceId;
        }

        // Check if instance exists
        var instance = _registry.GetInstance(instanceId);
        if (instance == null)
        {
            return new
            {
                content = new[] { new { type = "text", text = $"Instance {instanceId} is not registered or has timed out." } },
                isError = true
            };
        }

        // Dispatch command to the instance via command store
        string commandId = Guid.NewGuid().ToString("N")[..8];
        var cmd = new PendingCommand
        {
            Id = commandId,
            InstanceId = instanceId,
            Command = name,
            Args = args.HasValue ? JsonSerializer.Deserialize<object>(args.Value) : null,
            CreatedAt = DateTime.UtcNow,
        };
        _commandStore.AddCommand(cmd);

        // Wait for result (poll with timeout)
        var result = await _commandStore.WaitForResult(commandId, timeoutSeconds: 30);

        return result switch
        {
            { Success: true, Payload: var p } => new
            {
                content = new[] { new { type = "text", text = JsonSerializer.Serialize(p, new JsonSerializerOptions { WriteIndented = true }) } },
                isError = false
            },
            { Success: false, Error: var e } => new
            {
                content = new[] { new { type = "text", text = $"Command failed: {e}" } },
                isError = true
            },
            null => new
            {
                content = new[] { new { type = "text", text = "Command timed out waiting for add-in response." } },
                isError = true
            }
        };
    }

    /// <summary>
    /// Gets the registry for instance management endpoints.
    /// </summary>
    public static InstanceRegistry GetRegistry() => _registry;

    /// <summary>
    /// Gets the command store for dispatch/wait operations.
    /// </summary>
    public static CommandStore GetCommandStore() => _commandStore;
}
