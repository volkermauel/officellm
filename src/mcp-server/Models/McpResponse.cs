namespace OfficeMcpServer.Models;

/// <summary>
/// Standard MCP tool response envelope.
/// All tools return this structure for consistent parsing.
/// </summary>
public class McpResponse
{
    public bool Ok { get; set; }
    public string? ErrorCode { get; set; }
    public string? Message { get; set; }
    public string App { get; set; } = string.Empty;
    public string DocumentId { get; set; } = string.Empty;
    public object? Result { get; set; }
    public string[] Warnings { get; set; } = [];
    public bool RequiresConfirmation { get; set; }
    public string ConfirmationId { get; set; } = string.Empty;
    public string AuditId { get; set; } = string.Empty;

    /// <summary>
    /// Creates a successful response.
    /// </summary>
    public static McpResponse Success(string app, object? result = null, string[]? warnings = null)
        => new()
        {
            Ok = true,
            App = app,
            DocumentId = $"local-session:{app.ToLowerInvariant()}:active",
            Result = result,
            Warnings = warnings ?? [],
            RequiresConfirmation = false,
            AuditId = GenerateAuditId(),
        };

    /// <summary>
    /// Creates a confirmation-required response (for mutations).
    /// </summary>
    public static McpResponse RequiresConfirmationResponse(string app, object? result = null)
        => new()
        {
            Ok = true,
            App = app,
            DocumentId = $"local-session:{app.ToLowerInvariant()}:active",
            Result = result,
            RequiresConfirmation = true,
            ConfirmationId = Guid.NewGuid().ToString("N")[..8],
            AuditId = GenerateAuditId(),
        };

    /// <summary>
    /// Creates an error response.
    /// </summary>
    public static McpResponse Error(string errorCode, string message, bool recoverable = false)
        => new()
        {
            Ok = false,
            ErrorCode = errorCode,
            Message = message,
            Recoverable = recoverable,
            AuditId = GenerateAuditId(),
        };

    public bool Recoverable { get; set; }

    private static string GenerateAuditId()
        => DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-") + Guid.NewGuid().ToString("N")[..5];
}
