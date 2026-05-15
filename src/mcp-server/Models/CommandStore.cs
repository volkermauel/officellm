namespace OfficeMcpServer.Models;

/// <summary>
/// A pending command dispatched to an Office add-in instance.
/// </summary>
public class PendingCommand
{
    public string Id { get; set; } = string.Empty;
    public string InstanceId { get; set; } = string.Empty;
    public string Command { get; set; } = string.Empty;
    public object? Args { get; set; }
    public DateTime CreatedAt { get; set; }
    public string? ClaimedBy { get; set; }
    public bool Completed { get; set; }
    public bool Success { get; set; }
    public string? Error { get; set; }
    public object? Payload { get; set; }
    public DateTime? CompletedAt { get; set; }
}

/// <summary>
/// Thread-safe store for commands dispatched to add-in instances.
/// Supports queuing, claiming, result waiting, and cleanup.
/// </summary>
public class CommandStore
{
    private readonly Dictionary<string, PendingCommand> _commands = new();
    private readonly object _lock = new();

    /// <summary>
    /// Adds a new pending command.
    /// </summary>
    public void AddCommand(PendingCommand command)
    {
        lock (_lock)
        {
            _commands[command.Id] = command;
        }
    }

    /// <summary>
    /// Gets pending (unclaimed) commands for a specific instance.
    /// </summary>
    public List<PendingCommand> GetPendingCommands(string instanceId)
    {
        lock (_lock)
        {
            return _commands.Values
                .Where(c => c.InstanceId == instanceId && !c.Completed && string.IsNullOrEmpty(c.ClaimedBy))
                .ToList();
        }
    }

    /// <summary>
    /// Marks commands as claimed by an instance.
    /// </summary>
    public void MarkClaimed(string commandId, string instanceId)
    {
        lock (_lock)
        {
            if (_commands.TryGetValue(commandId, out var cmd) && string.IsNullOrEmpty(cmd.ClaimedBy))
            {
                cmd.ClaimedBy = instanceId;
            }
        }
    }

    /// <summary>
    /// Records the result of a completed command.
    /// </summary>
    public void CompleteCommand(string commandId, bool success, string error, object? payload)
    {
        lock (_lock)
        {
            if (_commands.TryGetValue(commandId, out var cmd))
            {
                cmd.Completed = true;
                cmd.Success = success;
                cmd.Error = error;
                cmd.Payload = payload;
                cmd.CompletedAt = DateTime.UtcNow;
            }
        }
    }

    /// <summary>
    /// Waits for a command result with a timeout.
    /// Polls every 500ms until complete or timeout.
    /// </summary>
    public async Task<PendingCommand?> WaitForResult(string commandId, int timeoutSeconds = 30)
    {
        var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);

        while (DateTime.UtcNow < deadline)
        {
            lock (_lock)
            {
                if (_commands.TryGetValue(commandId, out var cmd) && cmd.Completed)
                    return cmd;
            }
            await Task.Delay(500);
        }

        // Timeout — return null to signal timeout
        return null;
    }

    /// <summary>
    /// Cleans up old completed commands (older than 5 minutes).
    /// </summary>
    public void CleanupOldCommands()
    {
        lock (_lock)
        {
            var cutoff = DateTime.UtcNow.AddMinutes(-5);
            var toRemove = _commands.Values.Where(c => c.Completed && (c.CompletedAt ?? DateTime.MinValue) < cutoff).ToList();
            foreach (var cmd in toRemove)
                _commands.Remove(cmd.Id);
        }
    }
}
