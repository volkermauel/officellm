using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc.Testing;

namespace OfficeMcpServer.Tests;

/// <summary>
/// Integration tests using WebApplicationFactory.
/// Tests the full HTTP pipeline: routing, serialization, error handling.
/// </summary>
public class HttpEndpointTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly HttpClient _client;

    public HttpEndpointTests(WebApplicationFactory<Program> factory)
    {
        _client = factory.CreateClient();
        // Reset static state before each test class run
        Tools.McpToolEngine.ResetForTesting();
    }

    // --- Health ---

    [Fact]
    public async Task Health_ReturnsOk()
    {
        var response = await _client.GetAsync("/health");

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("ok", body.GetProperty("status").GetString());
        Assert.True(body.TryGetProperty("activeInstances", out _));
    }

    // --- Instances ---

    [Fact]
    public async Task Register_ReturnsInstanceId()
    {
        var response = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(body.TryGetProperty("instanceId", out var id));
        Assert.False(string.IsNullOrEmpty(id.GetString()));
    }

    [Fact]
    public async Task GetInstances_ReturnsRegisteredInstances()
    {
        // Register an instance first
        await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });

        var response = await _client.GetAsync("/instances");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var instances = body.GetProperty("instances");
        Assert.True(instances.GetArrayLength() >= 1);
    }

    [Fact]
    public async Task Heartbeat_ReturnsOk()
    {
        // Register first
        var reg = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });
        var regBody = await reg.Content.ReadFromJsonAsync<JsonElement>();
        var instanceId = regBody.GetProperty("instanceId").GetString();

        // Send heartbeat
        var response = await _client.PostAsJsonAsync($"/instances/{instanceId}/heartbeat", new
        {
            appName = "PowerPoint",
            documentName = "updated.pptx"
        });

        response.EnsureSuccessStatusCode();
    }

    [Fact]
    public async Task GetCommands_ReturnsEmpty_WhenNonePending()
    {
        var reg = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });
        var regBody = await reg.Content.ReadFromJsonAsync<JsonElement>();
        var instanceId = regBody.GetProperty("instanceId").GetString();

        var response = await _client.GetAsync($"/instances/{instanceId}/commands");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var commands = body.GetProperty("commands");
        Assert.Equal(0, commands.GetArrayLength());
    }

    [Fact]
    public async Task PostResult_ReturnsOk()
    {
        var reg = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });
        var regBody = await reg.Content.ReadFromJsonAsync<JsonElement>();
        var instanceId = regBody.GetProperty("instanceId").GetString();

        var response = await _client.PostAsJsonAsync($"/instances/{instanceId}/result", new
        {
            commandId = "test-cmd-1",
            success = true
        });

        response.EnsureSuccessStatusCode();
    }

    // --- OpenAPI ---

    [Fact]
    public async Task OpenApi_ReturnsValidSpec()
    {
        var response = await _client.GetAsync("/openapi.json");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("3.0.0", body.GetProperty("openapi").GetString());

        var info = body.GetProperty("info");
        Assert.Equal("Office LLM Harness API", info.GetProperty("title").GetString());

        // Must have paths
        var paths = body.GetProperty("paths");
        Assert.True(paths.EnumerateObject().Any(), "OpenAPI spec should have at least one path");

        // Must have /api/office_get_active_apps
        Assert.True(paths.TryGetProperty("/api/office_get_active_apps", out _),
            "Should have /api/office_get_active_apps path");
    }

    [Fact]
    public async Task OpenApi_ContainsAllToolEndpoints()
    {
        var response = await _client.GetAsync("/openapi.json");
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var paths = body.GetProperty("paths");

        var expectedPaths = new[]
        {
            "/api/office_get_active_apps",
            "/api/powerpoint_get_deck_outline",
            "/api/powerpoint_get_slide",
            "/api/powerpoint_update_shape_text",
            "/api/powerpoint_update_speaker_notes"
        };

        foreach (var expected in expectedPaths)
        {
            Assert.True(paths.TryGetProperty(expected, out var pathObj),
                $"Missing path: {expected}");
            Assert.True(pathObj.TryGetProperty("post", out var post), $"{expected} should have POST");
        }
    }

    [Fact]
    public async Task OpenApi_SpecHasServerUrl()
    {
        var response = await _client.GetAsync("/openapi.json");
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();

        var servers = body.GetProperty("servers");
        Assert.True(servers.GetArrayLength() > 0);

        var firstServer = servers[0];
        Assert.True(firstServer.TryGetProperty("url", out _));
    }

    // --- Swagger UI ---

    [Fact]
    public async Task Docs_ReturnsHtml()
    {
        var response = await _client.GetAsync("/docs");

        response.EnsureSuccessStatusCode();
        Assert.Equal("text/html", response.Content.Headers.ContentType?.MediaType);
        var html = await response.Content.ReadAsStringAsync();
        Assert.Contains("swagger-ui", html);
    }

    // --- REST API Bridge ---

    [Fact]
    public async Task ApiBridge_OfficeGetActiveApps_ReturnsSuccess()
    {
        var response = await _client.PostAsJsonAsync("/api/office_get_active_apps", new { });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(body.TryGetProperty("isError", out var isError));
        Assert.False(isError.GetBoolean());
    }

    [Fact]
    public async Task ApiBridge_UnknownTool_ReturnsError()
    {
        var response = await _client.PostAsJsonAsync("/api/nonexistent_tool", new { });

        // Should still return 200 with error payload (McpToolEngine returns error result)
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(body.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public async Task ApiBridge_InvalidJson_Returns400()
    {
        var content = new StringContent("not json", System.Text.Encoding.UTF8, "application/json");
        var response = await _client.PostAsync("/api/office_get_active_apps", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    // --- MCP Protocol ---

    [Fact]
    public async Task Mcp_Initialize_ReturnsCapabilities()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "initialize"
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("2.0", body.GetProperty("jsonrpc").GetString());

        var result = body.GetProperty("result");
        Assert.Equal("2024-11-05", result.GetProperty("protocolVersion").GetString());
        Assert.Equal("OfficeMcpServer", result.GetProperty("serverInfo").GetProperty("name").GetString());
    }

    [Fact]
    public async Task Mcp_ToolsList_ReturnsTools()
    {
        // Initialize first
        await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "initialize"
        });

        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 2,
            method = "tools/list"
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var tools = body.GetProperty("result").GetProperty("tools");
        Assert.True(tools.GetArrayLength() >= 5);

        // Verify tool schemas have proper properties
        var deckOutline = tools.EnumerateArray().First(t => t.GetProperty("name").GetString() == "powerpoint_get_deck_outline");
        var schemaProps = deckOutline.GetProperty("inputSchema").GetProperty("properties");
        Assert.True(schemaProps.ValueKind == JsonValueKind.Object, "properties should be an object");
        Assert.True(schemaProps.GetProperty("instanceId").GetProperty("type").GetString() == "string");

        var getSlide = tools.EnumerateArray().First(t => t.GetProperty("name").GetString() == "powerpoint_get_slide");
        var slideProps = getSlide.GetProperty("inputSchema").GetProperty("properties");
        Assert.True(slideProps.GetProperty("slideIndex").GetProperty("type").GetString() == "integer");
        var required = getSlide.GetProperty("inputSchema").GetProperty("required");
        Assert.Contains("instanceId", required.EnumerateArray().Select(r => r.GetString()));
        Assert.Contains("slideIndex", required.EnumerateArray().Select(r => r.GetString()));
    }

    [Fact]
    public async Task Mcp_ToolsCall_NoInstanceId_ReturnsError()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 3,
            method = "tools/call",
            @params = new
            {
                name = "powerpoint_get_deck_outline",
                arguments = new { }
            }
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var result = body.GetProperty("result");
        Assert.True(result.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public async Task Mcp_OfficeGetActiveApps_ReturnsSuccess()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 3,
            method = "tools/call",
            @params = new
            {
                name = "office_get_active_apps",
                arguments = new { }
            }
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var result = body.GetProperty("result");
        Assert.True(result.TryGetProperty("isError", out var isError));
        Assert.False(isError.GetBoolean());
    }

    [Fact]
    public async Task Mcp_EmptyBody_Returns400()
    {
        var content = new StringContent("", System.Text.Encoding.UTF8, "application/json");
        var response = await _client.PostAsync("/mcp", content);
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task Mcp_UnknownMethod_Returns404()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 99,
            method = "nonexistent/method"
        });

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }
}
