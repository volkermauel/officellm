using System.Net;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Http;
using OfficeMcpServer;
using OfficeMcpServer.Models;
using OfficeMcpServer.Tools;

// --- Global State ---
var bridgeState = new BridgeState();

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

// Session state (simple in-process for spike)
var sessionState = new { sessionId = Guid.NewGuid().ToString("N")[..12] };

Console.WriteLine($"Starting Office LLM Harness MCP Server on port {mcpPort}...");
Console.WriteLine("Bridge endpoint available at http://127.0.0.1:8765/");

// --- Bridge HTTP Server (port 8765) ---
var bridgeHandler = new BridgeHandler(bridgeState);
var bridgeTask = Task.Run(async () =>
{
    var bridgeServer = new HttpListener();
    bridgeServer.Prefixes.Add("http://127.0.0.1:8765/");
    bridgeServer.Start();
    Console.WriteLine("Bridge server listening on http://127.0.0.1:8765/");

    while (true)
    {
        try
        {
            var context = await bridgeServer.GetContextAsync();
            bridgeHandler.Handle(context);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Bridge error: {ex.Message}");
        }
    }
});

// --- MCP Protocol Endpoints (Streamable HTTP) ---

app.MapPost("/mcp", async (HttpContext context) =>
{
    var json = await new StreamReader(context.Request.Body).ReadToEndAsync();
    if (string.IsNullOrWhiteSpace(json))
    {
        return Results.BadRequest("Empty request body");
    }

    try
    {
        var message = JsonSerializer.Deserialize<JsonElement>(json);
        string method = message.GetProperty("method").GetString() ?? "";

        switch (method)
        {
            case "initialize":
                return Results.Json(new
                {
                    jsonrpc = "2.0",
                    id = message.GetProperty("id"),
                    result = new
                    {
                        protocolVersion = "2024-11-05",
                        serverInfo = new { name = "OfficeMcpServer", version = "0.1.0-spike" },
                        capabilities = new
                        {
                            tools = new { listChanged = true }
                        }
                    }
                }, new JsonSerializerOptions { WriteIndented = true });

            case "notifications/initialized":
                return Results.Ok(new { jsonrpc = "2.0", result = (object?)null });

            case "tools/list":
                object[] tools = GetToolDefinitions();
                return Results.Json(new
                {
                    jsonrpc = "2.0",
                    id = message.GetProperty("id"),
                    result = new { tools }
                }, new JsonSerializerOptions { WriteIndented = true });

            case "tools/call":
                string toolName = message.GetProperty("params").GetProperty("name").GetString() ?? "";
                JsonElement? toolArgs = null;
                if (message.GetProperty("params").TryGetProperty("input", out var input))
                    toolArgs = input;
                object result = await ExecuteTool(toolName, toolArgs);
                return Results.Json(new
                {
                    jsonrpc = "2.0",
                    id = message.GetProperty("id"),
                    result
                }, new JsonSerializerOptions { WriteIndented = true });

            default:
                return Results.Problem(detail: $"Method not found: {method}", statusCode: StatusCodes.Status404NotFound);
        }
    }
    catch (Exception ex)
    {
        return Results.Problem(detail: $"Parse error: {ex.Message}", statusCode: StatusCodes.Status400BadRequest);
    }
});

// Health check (non-MCP)
app.MapGet("/health", () => Results.Json(new
{
    status = "ok",
    sessionId = sessionState.sessionId,
    addInEndpoint = "http://127.0.0.1:8765",
    mcpPort = mcpPort
}));

app.Run($"http://127.0.0.1:{mcpPort}");

// --- Stdio JSON-RPC Server (for MCPo) ---
var stdioTask = Task.Run(async () =>
{
    Console.Error.WriteLine("stdio transport ready (MCPo compatible)");

    // Send initialized notification
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
                        jsonrpc = "2.0",
                        id = idEl,
                        result = new
                        {
                            protocolVersion = "2024-11-05",
                            serverInfo = new { name = "OfficeMcpServer", version = "0.1.0-spike" },
                            capabilities = new
                            {
                                tools = new { listChanged = true }
                            }
                        }
                    };
                    break;

                case "notifications/initialized":
                    // No response for notifications
                    break;

                case "tools/list":
                    object[] tools = GetToolDefinitions();
                    response = new
                    {
                        jsonrpc = "2.0",
                        id = idEl,
                        result = new { tools }
                    };
                    break;

                case "tools/call":
                    string toolName = message.GetProperty("params").GetProperty("name").GetString() ?? "";
                    JsonElement? toolArgs = null;
                    if (message.GetProperty("params").TryGetProperty("input", out var input))
                        toolArgs = input;
                    object result = await ExecuteTool(toolName, toolArgs);
                    response = new
                    {
                        jsonrpc = "2.0",
                        id = idEl,
                        result
                    };
                    break;

                default:
                    response = new
                    {
                        jsonrpc = "2.0",
                        id = idEl,
                        error = new { code = -32601, message = $"Method not found: {method}" }
                    };
                    break;
            }

            if (response != null)
                await WriteStdout(response);
        }
        catch (Exception ex)
        {
            // For errors, we can't reliably get the id, so omit it
            await WriteStdout(new
            {
                jsonrpc = "2.0",
                error = new { code = -32700, message = $"Parse error: {ex.Message}" }
            });
        }
    }
});

// --- Tool Definitions ---

static object[] GetToolDefinitions() => [
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
                ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                ["notes"] = new { type = "string", description = "Speaker notes text" }
            },
            required = new[] { "slideIndex", "notes" }
        }
    }
];

// --- Tool Execution ---

static async Task<object> ExecuteTool(string name, JsonElement? args)
{
    return name switch
    {
        "office_get_active_app" => await ExecuteGetActiveApp(),
        "powerpoint_get_deck_outline" => await OfficeTools.SendCommandToAddIn("PowerPoint", "getDeckOutline", args.HasValue ? JsonSerializer.Deserialize<object>(args.Value) : null),
        "powerpoint_get_slide" => await OfficeTools.SendCommandToAddIn("PowerPoint", "getSlide", args.HasValue ? JsonSerializer.Deserialize<object>(args.Value) : null),
        "powerpoint_update_shape_text" => await OfficeTools.SendCommandToAddIn("PowerPoint", "updateShapeText", args.HasValue ? JsonSerializer.Deserialize<object>(args.Value) : null),
        "powerpoint_update_speaker_notes" => await OfficeTools.SendCommandToAddIn("PowerPoint", "updateSpeakerNotes", args.HasValue ? JsonSerializer.Deserialize<object>(args.Value) : null),
        _ => throw new ArgumentException($"Unknown tool: {name}")
    };
}

static async Task<object> ExecuteGetActiveApp()
{
    var response = await OfficeTools.GetActiveApp();
    return new
    {
        content = new[] { new { type = "text", text = JsonSerializer.Serialize(response, new JsonSerializerOptions { WriteIndented = true }) } },
        isError = !response.Ok
    };
}

// --- Stdio Helpers ---

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

// --- Bridge Handler Class ---

class BridgeHandler
{
    private readonly BridgeState _state;

    public BridgeHandler(BridgeState state)
    {
        _state = state;
    }

    public void Handle(HttpListenerContext context)
    {
        var request = context.Request;
        var response = context.Response;

        // CORS
        response.Headers.Add("Access-Control-Allow-Origin", "*");
        response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
        response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");

        if (request.HttpMethod == "OPTIONS")
        {
            response.StatusCode = 204;
            response.Close();
            return;
        }

        string path = request.Url?.AbsolutePath ?? "/";

        if (request.HttpMethod == "POST" && path == "/command")
        {
            HandleCommandPost(request, response);
        }
        else if (request.HttpMethod == "GET" && path == "/result")
        {
            HandleResultGet(response);
        }
        else if (request.HttpMethod == "GET" && path == "/health")
        {
            HandleHealthGet(response);
        }
        else
        {
            response.StatusCode = 404;
            response.Close();
        }
    }

    private void HandleCommandPost(HttpListenerRequest request, HttpListenerResponse response)
    {
        using var reader = new StreamReader(request.InputStream);
        string body = reader.ReadToEnd();

        try
        {
            var json = JsonSerializer.Deserialize<JsonElement>(body);
            string command = json.GetProperty("command").GetString() ?? "unknown";
            JsonElement? argsEl = null;
            if (json.TryGetProperty("args", out var a))
                argsEl = a;

            object? argsObj = argsEl.HasValue ? JsonSerializer.Deserialize<object>(argsEl.Value) : null;

            lock (_state.Lock)
            {
                _state.Commands.Enqueue((command, argsObj, DateTime.UtcNow));
            }

            string responseData = JsonSerializer.Serialize(new { status = "queued", command });
            byte[] buffer = Encoding.UTF8.GetBytes(responseData);
            response.StatusCode = 200;
            response.ContentType = "application/json";
            response.ContentLength64 = buffer.Length;
            response.OutputStream.Write(buffer, 0, buffer.Length);
            response.Close();
        }
        catch
        {
            response.StatusCode = 400;
            response.ContentType = "application/json";
            byte[] err = Encoding.UTF8.GetBytes("{\"error\":\"Invalid JSON\"}");
            response.ContentLength64 = err.Length;
            response.OutputStream.Write(err, 0, err.Length);
            response.Close();
        }
    }

    private void HandleResultGet(HttpListenerResponse response)
    {
        string data;
        lock (_state.Lock)
        {
            data = _state.LastResult != null
                ? JsonSerializer.Serialize(_state.LastResult)
                : "{\"status\":\"idle\"}";
        }
        byte[] buffer2 = Encoding.UTF8.GetBytes(data);
        response.StatusCode = 200;
        response.ContentType = "application/json";
        response.ContentLength64 = buffer2.Length;
        response.OutputStream.Write(buffer2, 0, buffer2.Length);
        response.Close();
    }

    private void HandleHealthGet(HttpListenerResponse response)
    {
        int pending;
        lock (_state.Lock) pending = _state.Commands.Count;
        string h = JsonSerializer.Serialize(new { status = "ok", commandsPending = pending });
        byte[] hb = Encoding.UTF8.GetBytes(h);
        response.StatusCode = 200;
        response.ContentType = "application/json";
        response.ContentLength64 = hb.Length;
        response.OutputStream.Write(hb, 0, hb.Length);
        response.Close();
    }
}
