using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tests;

public class McpResponseTests
{
    [Fact]
    public void Success_CreatesOkResponse()
    {
        var response = McpResponse.Success("PowerPoint", new { slides = 5 });

        Assert.True(response.Ok);
        Assert.Equal("PowerPoint", response.App);
        Assert.NotNull(response.Result);
        Assert.Empty(response.Warnings);
        Assert.False(response.RequiresConfirmation);
        Assert.NotEmpty(response.AuditId);
    }

    [Fact]
    public void Success_WithWarnings()
    {
        var warnings = new[] { "Deprecated parameter used" };
        var response = McpResponse.Success("Word", null, warnings);

        Assert.True(response.Ok);
        Assert.Single(response.Warnings);
        Assert.Equal("Deprecated parameter used", response.Warnings[0]);
    }

    [Fact]
    public void Error_CreatesErrorResponse()
    {
        var response = McpResponse.Error("ADD_IN_UNREACHABLE", "Add-in not running", recoverable: true);

        Assert.False(response.Ok);
        Assert.Equal("ADD_IN_UNREACHABLE", response.ErrorCode);
        Assert.Equal("Add-in not running", response.Message);
        Assert.True(response.Recoverable);
        Assert.NotEmpty(response.AuditId);
    }

    [Fact]
    public void RequiresConfirmationResponse_SetsConfirmationFields()
    {
        var response = McpResponse.RequiresConfirmationResponse("PowerPoint", new { text = "new" });

        Assert.True(response.Ok);
        Assert.True(response.RequiresConfirmation);
        Assert.NotEmpty(response.ConfirmationId);
        Assert.Equal(8, response.ConfirmationId.Length); // Guid fragment
    }

    [Fact]
    public void Success_GeneratesDocumentId()
    {
        var response = McpResponse.Success("PowerPoint");

        Assert.Equal("local-session:powerpoint:active", response.DocumentId);
    }

    [Fact]
    public void AuditId_IsUnique()
    {
        var r1 = McpResponse.Success("PowerPoint");
        var r2 = McpResponse.Success("PowerPoint");

        Assert.NotEqual(r1.AuditId, r2.AuditId);
    }
}
