using System.Text.Json;
using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tools;

/// <summary>
/// Generic Office tools that work across all Office hosts.
/// Audit logging used by McpToolEngine.
/// </summary>
public static class OfficeTools
{
    private static readonly string AuditLogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "OfficeMcpServer", "audit.log");

    static OfficeTools()
    {
        // Ensure audit log directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(AuditLogPath)!);
    }

    public static void LogAudit(string tool, string? command, string status, string details)
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
            var json = JsonSerializer.Serialize(entry);
            File.AppendAllText(AuditLogPath, json + "\n");
        }
        catch
        {
            // Audit logging should never break the main flow
        }
    }
}
