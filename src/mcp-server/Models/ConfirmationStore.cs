namespace OfficeMcpServer.Models;

/// <summary>
/// Thread-safe store for pending confirmation tokens.
/// Used by mutation tools (update_shape_text, update_speaker_notes) to track
/// approved/rejected changes before they are applied.
/// </summary>
public class ConfirmationStore
{
    private readonly Dictionary<string, ConfirmationRequest> _confirmations = new();
    private readonly object _lock = new();

    /// <summary>
    /// Creates a new confirmation request and returns its token.
    /// </summary>
    public ConfirmationRequest Create(string toolName, string instanceId, int slideIndex, string? shapeId = null)
    {
        var request = new ConfirmationRequest
        {
            Token = Guid.NewGuid().ToString("N")[..8],
            ToolName = toolName,
            InstanceId = instanceId,
            SlideIndex = slideIndex,
            ShapeId = shapeId,
            Diff = new DiffPreview { OldText = "", NewText = "" }
        };

        lock (_lock)
        {
            _confirmations[request.Token] = request;
        }

        return request;
    }

    /// <summary>
    /// Validates a confirmation token and marks it as consumed.
    /// Returns true if the token is valid and hasn't been used yet.
    /// </summary>
    public bool ValidateToken(string token)
    {
        lock (_lock)
        {
            if (_confirmations.TryGetValue(token, out var req))
            {
                // Mark as consumed by removing it
                _confirmations.Remove(token);
                return true;
            }
            return false;
        }
    }

    /// <summary>
    /// Rejects a confirmation token (marks it as consumed without applying).
    /// </summary>
    public void RejectToken(string token)
    {
        lock (_lock)
        {
            _confirmations.Remove(token);
        }
    }

    /// <summary>
    /// Gets the confirmation request details for a token.
    /// </summary>
    public ConfirmationRequest? GetRequest(string token)
    {
        lock (_lock)
        {
            return _confirmations.TryGetValue(token, out var req) ? req : null;
        }
    }

    /// <summary>
    /// Gets all pending confirmation tokens for an instance.
    /// </summary>
    public List<string> GetPendingTokens(string instanceId)
    {
        lock (_lock)
        {
            return _confirmations.Values
                .Where(c => c.InstanceId == instanceId)
                .Select(c => c.Token)
                .ToList();
        }
    }
}
