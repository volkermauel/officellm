using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tests;

public class PowerPointModelsTests
{
    [Fact]
    public void SlideInfo_TextSummary_CombinesShapeTexts()
    {
        var slide = new SlideInfo
        {
            Index = 0,
            Title = "Introduction",
            Shapes =
            {
                new ShapeInfo { Id = "Title 1", Text = "Welcome" },
                new ShapeInfo { Id = "Content 1", Text = "Today's agenda" }
            }
        };

        Assert.Contains("Welcome", slide.TextSummary);
        Assert.Contains("Today's agenda", slide.TextSummary);
    }

    [Fact]
    public void DeckOutline_SetsCorrectCounts()
    {
        var outline = new DeckOutline
        {
            DocumentName = "test.pptx",
            TotalSlides = 10,
            Slides =
            {
                new SlideInfo { Index = 0 },
                new SlideInfo { Index = 1 }
            }
        };

        Assert.Equal("test.pptx", outline.DocumentName);
        Assert.Equal(10, outline.TotalSlides);
        Assert.Equal(2, outline.ReturnedSlides); // Should be set by caller
    }

    [Fact]
    public void DiffPreview_NoChanges_ShowsNoChangesMessage()
    {
        var diff = new DiffPreview
        {
            OldText = "Same text",
            NewText = "Same text"
        };

        Assert.Single(diff.UnifiedDiff);
        Assert.Equal("  (no changes)", diff.UnifiedDiff[0]);
    }

    [Fact]
    public void DiffPreview_LineLevelDiff_ShowsOldAndNew()
    {
        var diff = new DiffPreview
        {
            OldText = "Line 1\nOld line",
            NewText = "Line 1\nNew line"
        };

        var lines = diff.UnifiedDiff;
        Assert.Contains(lines, l => l.StartsWith("- Old line"));
        Assert.Contains(lines, l => l.StartsWith("+ New line"));
    }

    [Fact]
    public void ConfirmationRequest_GeneratesUniqueToken()
    {
        var c1 = new ConfirmationRequest();
        var c2 = new ConfirmationRequest();

        Assert.NotEqual(c1.Token, c2.Token);
        Assert.Equal(8, c1.Token.Length);
    }

    [Fact]
    public void ShapeInfo_IsTitle_FlaggedCorrectly()
    {
        var titleShape = new ShapeInfo { Id = "Title 1", Type = "title", IsTitle = true };
        var contentShape = new ShapeInfo { Id = "Content 1", Type = "content" };

        Assert.True(titleShape.IsTitle);
        Assert.False(contentShape.IsTitle);
    }
}
