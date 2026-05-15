using System.Text.Json;
using OfficeMcpServer.Tools;

namespace OfficeMcpServer.Tests;

public class McpToolEngineTests
{
	public McpToolEngineTests()
	{
		McpToolEngine.ResetForTesting();
	}
    [Fact]
    public void GetToolDefinitions_ReturnsExpectedTools()
    {
        var tools = McpToolEngine.GetToolDefinitions();

        Assert.NotNull(tools);
        Assert.True(tools.Length >= 5, $"Expected at least 5 tools, got {tools.Length}");
    }

    [Fact]
    public void GetToolDefinitions_ContainsExpectedToolNames()
    {
        var tools = McpToolEngine.GetToolDefinitions();
        // Serialize to JSON to access properties on anonymous types
        var names = tools.Select(t =>
        {
            var json = JsonSerializer.Serialize(t);
            var doc = JsonDocument.Parse(json);
            return doc.RootElement.GetProperty("name").GetString()!;
        }).ToList();

        Assert.Contains("office_get_active_apps", names);
        Assert.Contains("powerpoint_get_deck_outline", names);
        Assert.Contains("powerpoint_get_slide", names);
        Assert.Contains("powerpoint_update_shape_text", names);
        Assert.Contains("powerpoint_update_speaker_notes", names);
    }

    [Fact]
    public async Task ExecuteTool_NoInstances_ReturnsError()
    {
        // office_get_active_apps should succeed (returns empty list)
        var result = await McpToolEngine.ExecuteTool("office_get_active_apps", null);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("isError", out var isError));
        Assert.False(isError.GetBoolean());
    }

    [Fact]
    public async Task ExecuteTool_UnknownTool_ReturnsError()
    {
        var result = await McpToolEngine.ExecuteTool("nonexistent_tool", null);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }
}
