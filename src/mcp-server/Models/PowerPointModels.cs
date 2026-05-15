namespace OfficeMcpServer.Models;

/// <summary>
/// Represents a single shape's text content on a slide.
/// </summary>
public class ShapeInfo
{
    /// <summary>Shape name/ID within the slide.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Shape type (title, content, textbox, etc.).</summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>Text content of the shape.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Whether this shape is a title placeholder.</summary>
    public bool IsTitle { get; set; }

    /// <summary>X position of the shape (0-based).</summary>
    public int? Left { get; set; }

    /// <summary>Y position of the shape (0-based).</summary>
    public int? Top { get; set; }

    /// <summary>Width of the shape.</summary>
    public int? Width { get; set; }

    /// <summary>Height of the shape.</summary>
    public int? Height { get; set; }
}

/// <summary>
/// Represents a single slide in a PowerPoint deck.
/// </summary>
public class SlideInfo
{
    /// <summary>Zero-based slide index.</summary>
    public int Index { get; set; }

    /// <summary>Slide title (from first title placeholder, if any).</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>All shapes on the slide.</summary>
    public List<ShapeInfo> Shapes { get; set; } = new();

    /// <summary>Speaker notes for this slide (if requested).</summary>
    public string? SpeakerNotes { get; set; }

    /// <summary>Whether this slide is hidden.</summary>
    public bool IsHidden { get; set; }

    /// <summary>Summary of text content (first 200 chars of all shapes combined).</summary>
    public string TextSummary => string.Join(" ", Shapes.Select(s => s.Text).Take(3))
        .Replace("\n", " ").Trim();
}

/// <summary>
/// Full deck outline returned by powerpoint_get_deck_outline.
/// </summary>
public class DeckOutline
{
    /// <summary>Presentation document name.</summary>
    public string DocumentName { get; set; } = string.Empty;

    /// <summary>Total number of slides (before filtering).</summary>
    public int TotalSlides { get; set; }

    /// <summary>Number of slides in the response (after filtering). Auto-calculated from Slides.Count.</summary>
    public int ReturnedSlides => Slides.Count;

    /// <summary>List of slides in the outline.</summary>
    public List<SlideInfo> Slides { get; set; } = new();
}

/// <summary>
/// Before/after diff preview for a shape text change.
/// </summary>
public class DiffPreview
{
    /// <summary>Original text before the change.</summary>
    public string OldText { get; set; } = string.Empty;

    /// <summary>New text that would be applied.</summary>
    public string NewText { get; set; } = string.Empty;

    /// <summary>Character-level diff: lines starting with '-' are removed, '+' are added.</summary>
    public List<string> UnifiedDiff => ComputeUnifiedDiff();

    private List<string> ComputeUnifiedDiff()
    {
        var result = new List<string>();
        if (OldText == NewText)
        {
            result.Add("  (no changes)");
            return result;
        }

        // Simple line-level diff
        var oldLines = OldText.Split('\n');
        var newLines = NewText.Split('\n');

        foreach (var line in oldLines)
        {
            if (!newLines.Contains(line))
                result.Add($"- {line}");
        }
        foreach (var line in newLines)
        {
            if (!oldLines.Contains(line))
                result.Add($"+ {line}");
        }

        return result.Count > 0 ? result : new List<string> { "  (no visible changes)" };
    }
}

/// <summary>
/// Confirmation request for a mutation tool.
/// </summary>
public class ConfirmationRequest
{
    /// <summary>Unique confirmation token.</summary>
    public string Token { get; set; } = Guid.NewGuid().ToString("N")[..8];

    /// <summary>The diff preview to show in the task pane.</summary>
    public DiffPreview Diff { get; set; } = new();

    /// <summary>Tool name that requires confirmation.</summary>
    public string ToolName { get; set; } = string.Empty;

    /// <summary>Instance ID for the target add-in.</summary>
    public string InstanceId { get; set; } = string.Empty;

    /// <summary>Slide index involved in the change.</summary>
    public int SlideIndex { get; set; }

    /// <summary>Shape ID involved in the change.</summary>
    public string? ShapeId { get; set; }
}
