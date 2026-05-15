namespace OfficeMcpServer;

/// <summary>
/// Shared state for the bridge HTTP server.
/// Allows the HttpListener callback to access and modify queue/result data.
/// </summary>
public class BridgeState
{
    public Queue<(string Command, object? Args, DateTime Enqueued)> Commands { get; } = new();
    public object? LastResult { get; set; }
    public object Lock { get; } = new();
}
