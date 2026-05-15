using System.Text.Json;

namespace OfficeMcpServer.Models;

/// <summary>
/// Append-only JSONL audit log for all tool calls.
/// Each entry contains: timestamp, tool name, instanceId, inputs, confirmation status, outcome.
/// </summary>
public class AuditLog : IDisposable
{
    private readonly string _filePath;
    private readonly object _lock = new();
    private StreamWriter? _writer;
    private bool _disposed;

    public AuditLog(string? directoryPath = null)
    {
        var baseDir = directoryPath ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OfficeMcpServer");
        Directory.CreateDirectory(baseDir);
        _filePath = Path.Combine(baseDir, "audit.log");

        // Ensure file exists (create or truncate if first run)
        if (!File.Exists(_filePath))
            File.WriteAllText(_filePath, "");

        _writer = new StreamWriter(File.Open(_filePath, FileMode.Append, FileAccess.Write, FileShare.Read));
        _writer.AutoFlush = true;
    }

    /// <summary>
    /// Writes an audit entry to the JSONL file. Thread-safe (append-only).
    /// </summary>
    public void Log(AuditEntry entry)
    {
        lock (_lock)
        {
            if (_disposed) return;
            var json = JsonSerializer.Serialize(entry, new JsonSerializerOptions { WriteIndented = false });
            _writer!.WriteLine(json);
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _writer?.Dispose();
            _disposed = true;
        }
    }
}

/// <summary>
/// A single audit log entry in JSONL format.
/// </summary>
public class AuditEntry
{
    /// <summary>ISO 8601 UTC timestamp of the tool call.</summary>
    public string Timestamp { get; set; } = DateTime.UtcNow.ToString("o");

    /// <summary>Name of the MCP tool that was called.</summary>
    public string ToolName { get; set; } = string.Empty;

    /// <summary>Instance ID of the target Office add-in (if applicable).</summary>
    public string? InstanceId { get; set; }

    /// <summary>Raw inputs passed to the tool, serialized as JSON.</summary>
    public string Inputs { get; set; } = "{}";

    /// <summary>Whether this tool requires user confirmation (for mutations).</summary>
    public bool RequiresConfirmation { get; set; }

    /// <summary>Confirmation status: null=not required, "pending", "approved", "rejected".</summary>
    public string? ConfirmationStatus { get; set; }

    /// <summary>Tool outcome: "success", "error", "timeout", "rejected".</summary>
    public string Outcome { get; set; } = string.Empty;

    /// <summary>Error message if the tool call failed.</summary>
    public string? Error { get; set; }

    /// <summary>Unique identifier for this audit entry.</summary>
    public string AuditId { get; set; } = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-") + Guid.NewGuid().ToString("N")[..5];
}
