using System.Net.Http.Json;
using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tools;

/// <summary>
/// Generic Office tools that work across all Office hosts.
/// These communicate with the Office JS add-in via localhost HTTP.
/// </summary>
public static class OfficeTools
{
    private const string AddInBaseUrl = "http://127.0.0.1:8765";
    private static readonly HttpClient Http = new();
    private static readonly string AuditLogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "OfficeMcpServer", "audit.log");

    static OfficeTools()
    {
        // Ensure audit log directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(AuditLogPath)!);
    }

    /// <summary>
    /// office_get_active_app: Return active Office host, document name and selection metadata.
    /// </summary>
    public static async Task<McpResponse> GetActiveApp()
    {
        try
        {
            var response = await Http.GetAsync($"{AddInBaseUrl}/api/office/status");
            if (!response.IsSuccessStatusCode)
            {
                LogAudit("office_get_active_app", null, "add-in-unreachable", $"HTTP {response.StatusCode}");
                return McpResponse.Error(
                    "ADD_IN_UNREACHABLE",
                    "The Office add-in is not running. Please open PowerPoint and load the add-in task pane.",
                    recoverable: true);
            }

            var data = await response.Content.ReadFromJsonAsync<dynamic>();
            LogAudit("office_get_active_app", null, "success", data?.ToString() ?? "");
            return McpResponse.Success("Office", data);
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            LogAudit("office_get_active_app", null, "error", ex.Message);
            return McpResponse.Error(
                "ADD_IN_ERROR",
                $"Failed to communicate with Office add-in: {ex.Message}",
                recoverable: true);
        }
    }

    /// <summary>
    /// Sends a command to the Office add-in and returns the result.
    /// </summary>
    public static async Task<McpResponse> SendCommandToAddIn(string app, string command, object? args = null)
    {
        try
        {
            var payload = new { command, args };
            var response = await Http.PostAsJsonAsync($"{AddInBaseUrl}/api/command", payload);
            if (!response.IsSuccessStatusCode)
            {
                LogAudit(app.ToLowerInvariant(), command, "error", $"HTTP {response.StatusCode}");
                return McpResponse.Error(
                    "ADD_IN_ERROR",
                    $"Add-in returned HTTP {response.StatusCode}",
                    recoverable: true);
            }

            var result = await response.Content.ReadFromJsonAsync<dynamic>();
            var requiresConfirmation = result?["requiresConfirmation"]?.ToString() == "true";

            if (requiresConfirmation)
            {
                LogAudit(app.ToLowerInvariant(), command, "confirmation-required", result.ToString());
                return McpResponse.RequiresConfirmationResponse(app, result);
            }

            LogAudit(app.ToLowerInvariant(), command, "success", result?.ToString() ?? "");
            return McpResponse.Success(app, result);
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            LogAudit(app.ToLowerInvariant(), command, "error", ex.Message);
            return McpResponse.Error(
                "ADD_IN_ERROR",
                $"Failed to communicate with Office add-in: {ex.Message}",
                recoverable: true);
        }
    }

    private static void LogAudit(string tool, string? command, string status, string details)
    {
        try
        {
            var entry = new
            {
                timestamp = DateTime.UtcNow.ToString("o"),
                tool,
                command,
                status,
                details,
                auditId = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffff")
            };
            var json = System.Text.Json.JsonSerializer.Serialize(entry);
            File.AppendAllText(AuditLogPath, json + "\n");
        }
        catch
        {
            // Audit logging should never break the main flow
        }
    }
}
