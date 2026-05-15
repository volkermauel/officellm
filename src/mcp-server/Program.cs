using System.Text;
using System.Text.Json;
using OfficeMcpServer;
using OfficeMcpServer.Tools;

// --- MCP configuration ---
var mcpPort = args.FirstOrDefault(a => !a.StartsWith("--")) is string portArg && int.TryParse(portArg, out var p) ? p : 3000;

Console.WriteLine($"Starting Office LLM Harness MCP Server on port {mcpPort}...");

var app = AppBuilder.Create(args, mcpPort);

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
            var registry = McpToolEngine.GetRegistry();
            var commandStore = McpToolEngine.GetCommandStore();

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

// Start the HTTP server (blocks until shutdown)
app.Run();

// Required for WebApplicationFactory integration testing
public partial class Program { }
