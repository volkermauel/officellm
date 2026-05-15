using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Http;
using OfficeMcpServer.Models;
using OfficeMcpServer.Tools;

// --- Global State ---
var registry = McpToolEngine.GetRegistry();
var commandStore = McpToolEngine.GetCommandStore();

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddCors(options =>
{
    options.AddPolicy("LocalOnly", policy =>
    {
        policy.WithOrigins("http://127.0.0.1:*", "https://127.0.0.1:*")
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();
app.UseCors("LocalOnly");

// MCP configuration
var mcpPort = args.Length > 0 ? int.Parse(args[0]) : 3000;

Console.WriteLine($"Starting Office LLM Harness MCP Server on port {mcpPort}...");

// Start cleanup timer (runs every 30s)
_ = Task.Run(async () =>
{
    while (true)
    {
        await Task.Delay(30000);
        registry.CleanupTimedOut();
    }
});

// ============================================================
// INSTANCE MANAGEMENT ENDPOINTS (for Office Add-ins)
// ============================================================

app.MapPost("/instances/register", async (HttpContext context) =>
{
    var json = await new StreamReader(context.Request.Body).ReadToEndAsync();
    var data = JsonSerializer.Deserialize<JsonElement>(json);
    string appName = data.GetProperty("appName").GetString() ?? "Unknown";
    string documentName = data.TryGetProperty("documentName", out var dn) ? dn.GetString() ?? "" : "";

    string instanceId = registry.RegisterInstance(appName, documentName);
    Console.WriteLine($"Registered instance: {instanceId} ({appName} - {documentName})");
    return Results.Json(new { instanceId, appName, documentName });
});

app.MapPost("/instances/{instanceId}/heartbeat", async (HttpContext context, string instanceId) =>
{
    var json = await new StreamReader(context.Request.Body).ReadToEndAsync();
    var data = JsonSerializer.Deserialize<JsonElement>(json);
    string? appName = data.TryGetProperty("appName", out var a) ? a.GetString() : null;
    string? documentName = data.TryGetProperty("documentName", out var d) ? d.GetString() : null;
    registry.UpdateHeartbeat(instanceId, appName, documentName);
    return Results.Ok(new { status = "ok" });
});

app.MapGet("/instances", () =>
{
    var instances = registry.GetActiveInstances().Select(i => new
    {
        i.InstanceId, i.AppName, i.DocumentName, i.IsAlive,
        RegisteredAt = i.RegisteredAt.ToString("o"),
    });
    return Results.Json(new { instances });
});

app.MapGet("/instances/{instanceId}/commands", async (string instanceId) =>
{
    var commands = commandStore.GetPendingCommands(instanceId);
    if (commands.Any())
        foreach (var cmd in commands)
            commandStore.MarkClaimed(cmd.Id, instanceId);
    return Results.Json(new { commands });
});

app.MapPost("/instances/{instanceId}/result", async (HttpContext context, string instanceId) =>
{
    var json = await new StreamReader(context.Request.Body).ReadToEndAsync();
    var data = JsonSerializer.Deserialize<JsonElement>(json);
    string commandId = data.GetProperty("commandId").GetString() ?? "";
    bool success = data.TryGetProperty("success", out var s) && s.GetBoolean();
    string? error = data.TryGetProperty("error", out var e) ? e.GetString() : null;
    object? payload = data.TryGetProperty("payload", out var p) ? p.Deserialize<object>() : null;
    Console.WriteLine($"Instance {instanceId} completed command {commandId}: success={success}");
    commandStore.CompleteCommand(commandId, success, error ?? "", payload);
    return Results.Ok(new { status = "ok" });
});

// ============================================================
// MCP PROTOCOL ENDPOINTS (Streamable HTTP)
// ============================================================

app.MapPost("/mcp", async (HttpContext context) =>
{
    var json = await new StreamReader(context.Request.Body).ReadToEndAsync();
    if (string.IsNullOrWhiteSpace(json))
        return Results.BadRequest("Empty request body");

    try
    {
        var message = JsonSerializer.Deserialize<JsonElement>(json);
        string method = message.GetProperty("method").GetString() ?? "";

        switch (method)
        {
            case "initialize":
                return Results.Json(new
                {
                    jsonrpc = "2.0", id = message.GetProperty("id"),
                    result = new
                    {
                        protocolVersion = "2024-11-05",
                        serverInfo = new { name = "OfficeMcpServer", version = "0.1.0-spike" },
                        capabilities = new { tools = new { listChanged = true } }
                    }
                }, new JsonSerializerOptions { WriteIndented = true });

            case "notifications/initialized":
                return Results.Ok(new { jsonrpc = "2.0", result = (object?)null });

            case "tools/list":
                var tools = McpToolEngine.GetToolDefinitions();
                return Results.Json(new
                { jsonrpc = "2.0", id = message.GetProperty("id"), result = new { tools } },
                new JsonSerializerOptions { WriteIndented = true });

            case "tools/call":
                var toolName = message.GetProperty("params").GetProperty("name").GetString() ?? "";
                JsonElement? toolArgs = null;
                if (message.GetProperty("params").TryGetProperty("input", out var input))
                    toolArgs = input;
                var result = await McpToolEngine.ExecuteTool(toolName, toolArgs);
                return Results.Json(new
                { jsonrpc = "2.0", id = message.GetProperty("id"), result },
                new JsonSerializerOptions { WriteIndented = true });

            default:
                return Results.Problem(detail: $"Method not found: {method}", statusCode: StatusCodes.Status404NotFound);
        }
    }
    catch (Exception ex)
    {
        return Results.Problem(detail: $"Parse error: {ex.Message}", statusCode: StatusCodes.Status400BadRequest);
    }
});

app.MapGet("/health", () =>
{
    var activeCount = registry.GetActiveInstances().Count;
    return Results.Json(new { status = "ok", activeInstances = activeCount, mcpPort });
});

// ============================================================
// STDIO JSON-RPC SERVER (for MCPo)
// ============================================================

var stdioTask = Task.Run(async () =>
{
    Console.Error.WriteLine("stdio transport ready (MCPo compatible)");
    await WriteStdout(new { jsonrpc = "2.0", @event = "notifications/initialized" });

    await foreach (var line in ReadStdinLines())
    {
        if (string.IsNullOrWhiteSpace(line)) continue;

        try
        {
            var message = JsonSerializer.Deserialize<JsonElement>(line);
            string method = message.GetProperty("method").GetString() ?? "";
            JsonElement? idEl = null;
            if (message.TryGetProperty("id", out var id))
                idEl = id;

            object? response = null;

            switch (method)
            {
                case "initialize":
                    response = new
                    {
                        jsonrpc = "2.0", id = idEl,
                        result = new
                        {
                            protocolVersion = "2024-11-05",
                            serverInfo = new { name = "OfficeMcpServer", version = "0.1.0-spike" },
                            capabilities = new { tools = new { listChanged = true } }
                        }
                    };
                    break;

                case "notifications/initialized":
                    break;

                case "tools/list":
                    var tools = McpToolEngine.GetToolDefinitions();
                    response = new { jsonrpc = "2.0", id = idEl, result = new { tools } };
                    break;

                case "tools/call":
                    var tName = message.GetProperty("params").GetProperty("name").GetString() ?? "";
                    JsonElement? tArgs = null;
                    if (message.GetProperty("params").TryGetProperty("input", out var tInput))
                        tArgs = tInput;
                    var tResult = await McpToolEngine.ExecuteTool(tName, tArgs);
                    response = new { jsonrpc = "2.0", id = idEl, result = tResult };
                    break;

                default:
                    response = new
                    {
                        jsonrpc = "2.0", id = idEl,
                        error = new { code = -32601, message = $"Method not found: {method}" }
                    };
                    break;
            }

            if (response != null)
                await WriteStdout(response);
        }
        catch (Exception ex)
        {
            await WriteStdout(new
            {
                jsonrpc = "2.0",
                error = new { code = -32700, message = $"Parse error: {ex.Message}" }
            });
        }
    }
});

// ============================================================
// STDIO HELPERS
// ============================================================

static async IAsyncEnumerable<string> ReadStdinLines()
{
    using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8, leaveOpen: true);
    while (!reader.EndOfStream)
    {
        string? line = await reader.ReadLineAsync();
        if (line != null)
            yield return line;
    }
}

static async Task WriteStdout(object obj)
{
    string json = JsonSerializer.Serialize(obj, new JsonSerializerOptions { WriteIndented = false });
    byte[] data = Encoding.UTF8.GetBytes(json + "\n");
    await Console.OpenStandardOutput().WriteAsync(data, 0, data.Length);
    await Console.OpenStandardOutput().FlushAsync();
}
