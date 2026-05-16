namespace OfficeMcpServer.Models;

/// <summary>
/// Represents a registered Office application instance (e.g., one PowerPoint session).
/// Each add-in registers itself with a unique instance ID.
/// </summary>
public class OfficeInstance
{
    public string InstanceId { get; set; } = string.Empty;
    public string AppName { get; set; } = string.Empty;
    public string DocumentName { get; set; } = string.Empty;
    public DateTime RegisteredAt { get; set; } = DateTime.UtcNow;
    public DateTime LastHeartbeat { get; set; } = DateTime.UtcNow;
    public bool IsAlive { get; set; } = true;

    /// <summary>
    /// Updates the heartbeat timestamp.
    /// </summary>
    public void RefreshHeartbeat()
    {
        LastHeartbeat = DateTime.UtcNow;
        IsAlive = true;
    }

    /// <summary>
    /// Checks if the instance has timed out (no heartbeat for 60 seconds).
    /// </summary>
    public bool HasTimedOut(int timeoutSeconds = 60)
    {
        return (DateTime.UtcNow - LastHeartbeat).TotalSeconds > timeoutSeconds;
    }
}

/// <summary>
/// Thread-safe registry of active Office application instances.
/// The MCP server uses this to route tool calls to the correct add-in.
/// </summary>
public class InstanceRegistry
{
    private readonly Dictionary<string, OfficeInstance> _instances = new();
    private int _nextInstanceId = 1;
    private readonly object _lock = new();

    private static string GetHostPrefix(string appName)
    {
        // Normalize appName to host prefix
        var lower = appName.ToLowerInvariant();
        if (lower.Contains("word")) return "word_";
        if (lower.Contains("excel")) return "excel_";
        if (lower.Contains("outlook")) return "outlook_";
        if (lower.Contains("powerpoint") || lower.Contains("ppt")) return "powerpoint_";
        return "office_";
    }

    /// <summary>
    /// Registers a new Office instance and returns its ID.
    /// </summary>
    public string RegisterInstance(string appName, string documentName)
    {
        lock (_lock)
        {
            string instanceId = $"{GetHostPrefix(appName)}{_nextInstanceId++}";

            // If this app name already exists, reuse it (update in place)
            if (_instances.ContainsKey(instanceId))
            {
                var existing = _instances[instanceId];
                existing.AppName = appName;
                existing.DocumentName = documentName;
                existing.RefreshHeartbeat();
                return instanceId;
            }

            _instances[instanceId] = new OfficeInstance
            {
                InstanceId = instanceId,
                AppName = appName,
                DocumentName = documentName,
            };

            return instanceId;
        }
    }

    /// <summary>
    /// Updates an existing instance's heartbeat and info.
    /// </summary>
    public void UpdateHeartbeat(string instanceId, string? appName = null, string? documentName = null)
    {
        lock (_lock)
        {
            if (_instances.TryGetValue(instanceId, out var instance))
            {
                instance.RefreshHeartbeat();
                if (appName != null) instance.AppName = appName;
                if (documentName != null) instance.DocumentName = documentName;
            }
        }
    }

    /// <summary>
    /// Gets all active (non-timed-out) instances.
    /// </summary>
    public List<OfficeInstance> GetActiveInstances()
    {
        lock (_lock)
        {
            return _instances.Values.Where(i => !i.HasTimedOut()).ToList();
        }
    }

    /// <summary>
    /// Gets a specific instance by ID, or null if not found.
    /// </summary>
    public OfficeInstance? GetInstance(string instanceId)
    {
        lock (_lock)
        {
            return _instances.TryGetValue(instanceId, out var instance) && !instance.HasTimedOut()
                ? instance
                : null;
        }
    }

    /// <summary>
    /// Removes timed-out instances (no heartbeat for 60 seconds).
    /// </summary>
    public void CleanupTimedOut()
    {
        lock (_lock)
        {
            var timedOut = _instances.Values.Where(i => i.HasTimedOut(60)).ToList();
            foreach (var instance in timedOut)
            {
                _instances.Remove(instance.InstanceId);
                Console.WriteLine($"Instance {instance.InstanceId} timed out and removed");
            }
        }
    }
}
