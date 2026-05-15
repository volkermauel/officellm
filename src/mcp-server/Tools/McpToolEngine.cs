using System.Text.Json;
using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tools;

/// <summary>
/// Handles MCP tool definition generation and command dispatch.
/// Uses instance-aware routing with a fixed tool list.
/// </summary>
public static class McpToolEngine
{
    private static InstanceRegistry _registry = new();
    private static CommandStore _commandStore = new();
    private static ConfirmationStore _confirmationStore = new();
    private static AuditLog _auditLog = new();

    /// <summary>
    /// Gets all MCP tool definitions (fixed list, instance selected via parameter).
    /// </summary>
    public static object[] GetToolDefinitions() => [
        new
        {
            name = "office_get_active_apps",
            description = "Returns a list of all registered Office instances (PowerPoint, Word, Excel, Outlook) with their document names. Use this to discover which documents are open and let the user choose which one to work with.",
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
            description = "Returns the full slide deck outline with titles and text content for each slide. Use office_get_active_apps first to find the right instanceId.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "Target instance ID (e.g. 'powerpoint_1'). Omit to use the most recently registered instance." },
                    ["includeSpeakerNotes"] = new { type = "boolean", description = "Include speaker notes in the outline", @default = false },
                    ["includeHiddenSlides"] = new { type = "boolean", description = "Include hidden slides in the outline", @default = false }
                },
                required = Array.Empty<string>()
            }
        },
        new
        {
            name = "powerpoint_get_slide",
            description = "Returns all shapes with their text content for a single slide. IMPORTANT: You MUST provide slideIndex (zero-based). Use powerpoint_get_deck_outline first to see all slide numbers.",
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
                    ["text"] = new { type = "string", description = "New text content for the shape" },
                    ["confirmationToken"] = new { type = "string", description = "Confirmation token from pending change (required for mutations)" }
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
                    ["notes"] = new { type = "string", description = "Speaker notes text" },
                    ["confirmationToken"] = new { type = "string", description = "Confirmation token from pending change (required for mutations)" }
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

        // Parse arguments
        var inputs = args.HasValue ? JsonSerializer.Serialize(args.Value) : "{}";
        var confirmationToken = (args.HasValue && args.Value.TryGetProperty("confirmationToken", out var ct)) ? ct.GetString() : null;

        // Route to appropriate handler
        switch (name)
        {
            case "office_get_active_apps":
                return HandleGetActiveApps();

            case "office_get_active_app":
                return HandleGetActiveApps();

            case "powerpoint_get_deck_outline":
                return await HandleGetDeckOutline(instanceId, args, inputs);

            case "powerpoint_get_slide":
                return await HandleGetSlide(instanceId, args, inputs);

            case "powerpoint_update_shape_text":
                return await HandleUpdateShapeText(instanceId, args, inputs, confirmationToken);

            case "powerpoint_update_speaker_notes":
                return await HandleUpdateSpeakerNotes(instanceId, args, inputs, confirmationToken);

            default:
                return new
                {
                    content = new[] { new { type = "text", text = $"Unknown tool: {name}" } },
                    isError = true
                };
        }
    }

    private static object HandleGetActiveApps()
    {
        var instances = _registry.GetActiveInstances();

        _auditLog.Log(new AuditEntry
        {
            ToolName = "office_get_active_apps",
            InstanceId = "",
            Inputs = "",
            Outcome = "success"
        });

        var appList = instances.Select(i => new
        {
            instanceId = i.InstanceId,
            appName = i.AppName,
            documentName = i.DocumentName
        }).ToList();

        return new
        {
            content = new[] { new { type = "text", text = JsonSerializer.Serialize(new { apps = appList, total = appList.Count }, new JsonSerializerOptions { WriteIndented = true }) } },
            isError = false
        };
    }

    private static async Task<object> HandleGetDeckOutline(string instanceId, JsonElement? args, string inputs)
    {
        var includeNotes = args.HasValue && args.Value.TryGetProperty("includeSpeakerNotes", out var in_) && in_.GetBoolean();
        var includeHidden = args.HasValue && args.Value.TryGetProperty("includeHiddenSlides", out var ih) && ih.GetBoolean();

        // Dispatch command to add-in to get deck outline via Office JS API
        string commandId = Guid.NewGuid().ToString("N")[..8];
        var cmd = new PendingCommand
        {
            Id = commandId,
            InstanceId = instanceId,
            Command = "powerpoint_get_deck_outline",
            Args = new { includeSpeakerNotes = includeNotes, includeHiddenSlides = includeHidden },
            CreatedAt = DateTime.UtcNow,
        };
        _commandStore.AddCommand(cmd);

        _auditLog.Log(new AuditEntry
        {
            ToolName = "powerpoint_get_deck_outline",
            InstanceId = instanceId,
            Inputs = inputs,
            Outcome = "pending"
        });

        var result = await _commandStore.WaitForResult(commandId, timeoutSeconds: 30);
        return BuildToolResult(result, "powerpoint_get_deck_outline", instanceId, inputs);
    }

    private static async Task<object> HandleGetSlide(string instanceId, JsonElement? args, string inputs)
    {
        if (!args.HasValue || !args.Value.TryGetProperty("slideIndex", out var si) || !si.TryGetInt32(out var slideIndex))
        {
            return new
            {
                content = new[] { new { type = "text", text = "Missing required parameter: slideIndex" } },
                isError = true
            };
        }

        string commandId = Guid.NewGuid().ToString("N")[..8];
        var cmd = new PendingCommand
        {
            Id = commandId,
            InstanceId = instanceId,
            Command = "powerpoint_get_slide",
            Args = new { slideIndex },
            CreatedAt = DateTime.UtcNow,
        };
        _commandStore.AddCommand(cmd);

        _auditLog.Log(new AuditEntry
        {
            ToolName = "powerpoint_get_slide",
            InstanceId = instanceId,
            Inputs = inputs,
            Outcome = "pending"
        });

        var result = await _commandStore.WaitForResult(commandId, timeoutSeconds: 30);
        return BuildToolResult(result, "powerpoint_get_slide", instanceId, inputs);
    }

    private static async Task<object> HandleUpdateShapeText(string instanceId, JsonElement? args, string inputs, string? confirmationToken)
    {
        if (!args.HasValue || !args.Value.TryGetProperty("slideIndex", out var si) || !si.TryGetInt32(out var slideIndex))
        {
            return new
            {
                content = new[] { new { type = "text", text = "Missing required parameter: slideIndex" } },
                isError = true
            };
        }

        if (!args.HasValue || !args.Value.TryGetProperty("shapeId", out var shapeIdEl) || string.IsNullOrEmpty(shapeIdEl.GetString()))
        {
            return new
            {
                content = new[] { new { type = "text", text = "Missing required parameter: shapeId" } },
                isError = true
            };
        }

        if (!args.HasValue || !args.Value.TryGetProperty("text", out var textEl) || string.IsNullOrEmpty(textEl.GetString()))
        {
            return new
            {
                content = new[] { new { type = "text", text = "Missing required parameter: text" } },
                isError = true
            };
        }

        var newText = textEl.GetString()!;
        var shapeId = shapeIdEl.GetString()!;

        // First call (no token) → create confirmation request with diff preview
        if (string.IsNullOrEmpty(confirmationToken))
        {
            var confirmation = _confirmationStore.Create(
                "powerpoint_update_shape_text",
                instanceId,
                slideIndex,
                shapeId);

            // Get current text from add-in for diff (stub - would need Office JS API call)
            var diff = new DiffPreview
            {
                OldText = "[current text]",
                NewText = newText
            };

            confirmation.Diff = diff;

            _auditLog.Log(new AuditEntry
            {
                ToolName = "powerpoint_update_shape_text",
                InstanceId = instanceId,
                Inputs = inputs,
                RequiresConfirmation = true,
                ConfirmationStatus = "pending",
                Outcome = "confirmation_required"
            });

            return new
            {
                requiresConfirmation = true,
                confirmationToken = confirmation.Token,
                diff = diff,
                slideIndex = slideIndex,
                shapeId = shapeId,
                newText = newText,
                message = "Please review the diff and approve in the task pane"
            };
        }

        // Second call (with token) → validate and apply
        if (!_confirmationStore.ValidateToken(confirmationToken))
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = "powerpoint_update_shape_text",
                InstanceId = instanceId,
                Inputs = inputs,
                RequiresConfirmation = true,
                ConfirmationStatus = "rejected",
                Outcome = "error",
                Error = "Invalid or expired confirmation token"
            });

            return new
            {
                content = new[] { new { type = "text", text = "Invalid or expired confirmation token" } },
                isError = true
            };
        }

        // Dispatch command to add-in to apply the change
        string commandId = Guid.NewGuid().ToString("N")[..8];
        var cmd = new PendingCommand
        {
            Id = commandId,
            InstanceId = instanceId,
            Command = "powerpoint_update_shape_text",
            Args = new { slideIndex, shapeId, text = newText, confirmationToken },
            CreatedAt = DateTime.UtcNow,
        };
        _commandStore.AddCommand(cmd);

        _auditLog.Log(new AuditEntry
        {
            ToolName = "powerpoint_update_shape_text",
            InstanceId = instanceId,
            Inputs = inputs,
            RequiresConfirmation = true,
            ConfirmationStatus = "approved",
            Outcome = "pending"
        });

        var result = await _commandStore.WaitForResult(commandId, timeoutSeconds: 30);
        return BuildToolResult(result, "powerpoint_update_shape_text", instanceId, inputs);
    }

    private static async Task<object> HandleUpdateSpeakerNotes(string instanceId, JsonElement? args, string inputs, string? confirmationToken)
    {
        if (!args.HasValue || !args.Value.TryGetProperty("slideIndex", out var si) || !si.TryGetInt32(out var slideIndex))
        {
            return new
            {
                content = new[] { new { type = "text", text = "Missing required parameter: slideIndex" } },
                isError = true
            };
        }

        if (!args.HasValue || !args.Value.TryGetProperty("notes", out var notesEl) || string.IsNullOrEmpty(notesEl.GetString()))
        {
            return new
            {
                content = new[] { new { type = "text", text = "Missing required parameter: notes" } },
                isError = true
            };
        }

        var newNotes = notesEl.GetString()!;

        // First call (no token) → create confirmation request
        if (string.IsNullOrEmpty(confirmationToken))
        {
            var confirmation = _confirmationStore.Create(
                "powerpoint_update_speaker_notes",
                instanceId,
                slideIndex);

            var diff = new DiffPreview
            {
                OldText = "[current notes]",
                NewText = newNotes
            };

            confirmation.Diff = diff;

            _auditLog.Log(new AuditEntry
            {
                ToolName = "powerpoint_update_speaker_notes",
                InstanceId = instanceId,
                Inputs = inputs,
                RequiresConfirmation = true,
                ConfirmationStatus = "pending",
                Outcome = "confirmation_required"
            });

            return new
            {
                requiresConfirmation = true,
                confirmationToken = confirmation.Token,
                diff = diff,
                slideIndex = slideIndex,
                newNotes = newNotes,
                message = "Please review the notes and approve in the task pane"
            };
        }

        // Second call (with token) → validate and apply
        if (!_confirmationStore.ValidateToken(confirmationToken))
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = "powerpoint_update_speaker_notes",
                InstanceId = instanceId,
                Inputs = inputs,
                RequiresConfirmation = true,
                ConfirmationStatus = "rejected",
                Outcome = "error",
                Error = "Invalid or expired confirmation token"
            });

            return new
            {
                content = new[] { new { type = "text", text = "Invalid or expired confirmation token" } },
                isError = true
            };
        }

        // Dispatch command to add-in
        string commandId = Guid.NewGuid().ToString("N")[..8];
        var cmd = new PendingCommand
        {
            Id = commandId,
            InstanceId = instanceId,
            Command = "powerpoint_update_speaker_notes",
            Args = new { slideIndex, notes = newNotes, confirmationToken },
            CreatedAt = DateTime.UtcNow,
        };
        _commandStore.AddCommand(cmd);

        _auditLog.Log(new AuditEntry
        {
            ToolName = "powerpoint_update_speaker_notes",
            InstanceId = instanceId,
            Inputs = inputs,
            RequiresConfirmation = true,
            ConfirmationStatus = "approved",
            Outcome = "pending"
        });

        var result = await _commandStore.WaitForResult(commandId, timeoutSeconds: 30);
        return BuildToolResult(result, "powerpoint_update_speaker_notes", instanceId, inputs);
    }

    private static object BuildToolResult(PendingCommand? result, string toolName, string instanceId, string inputs)
    {
        if (result == null)
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = toolName,
                InstanceId = instanceId,
                Inputs = inputs,
                Outcome = "timeout"
            });

            return new
            {
                content = new[] { new { type = "text", text = "Command timed out waiting for add-in response." } },
                isError = true
            };
        }

        if (result.Success)
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = toolName,
                InstanceId = instanceId,
                Inputs = inputs,
                Outcome = "success"
            });

            return new
            {
                content = new[] { new { type = "text", text = JsonSerializer.Serialize(result.Payload, new JsonSerializerOptions { WriteIndented = true }) } },
                isError = false
            };
        }

        _auditLog.Log(new AuditEntry
        {
            ToolName = toolName,
            InstanceId = instanceId,
            Inputs = inputs,
            Outcome = "error",
            Error = result.Error
        });

        return new
        {
            content = new[] { new { type = "text", text = $"Command failed: {result.Error}" } },
            isError = true
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

    /// <summary>
    /// Gets the confirmation store for mutation tool confirmations.
    /// </summary>
    public static ConfirmationStore GetConfirmationStore() => _confirmationStore;

    /// <summary>
    /// Resets all static state. Used by tests to isolate test runs.
    /// </summary>
    public static void ResetForTesting()
    {
        _registry = new InstanceRegistry();
        _commandStore = new CommandStore();
        _confirmationStore = new ConfirmationStore();
        _auditLog = new AuditLog(Path.Combine(Path.GetTempPath(), $"audit-test-{Guid.NewGuid()}").TrimEnd('/'));
    }
}
