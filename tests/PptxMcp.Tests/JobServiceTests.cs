using PptxMcp.Jobs;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class JobServiceTests
{
    [Theory]
    [InlineData("preview/slide-1.png", 1)]
    [InlineData("preview/slide-01.png", 1)]
    [InlineData("preview/slide-50.png", 50)]
    public void ParsesPaddedAndUnpaddedPreviewSlideNumbers(string fileName, int expected)
    {
        Assert.Equal(expected, JobService.TryGetPreviewSlideNumber(fileName));
    }

    [Theory]
    [InlineData("slide-01.png")]
    [InlineData("preview/slide-one.png")]
    [InlineData("preview/slide-01.jpg")]
    public void RejectsUnexpectedPreviewArtifactNames(string fileName)
    {
        Assert.Null(JobService.TryGetPreviewSlideNumber(fileName));
    }

    [Fact]
    public void AppliesOnlyRequestedDeckRevisionsAndPreservesLayouts()
    {
        var slides = new[]
        {
            new DeckSlideSpec("layout-1", [new DeckField("Original 1", ShapeId: 2)]),
            new DeckSlideSpec("layout-2", [new DeckField("Original 2", ShapeId: 3)]),
            new DeckSlideSpec("layout-3", [new DeckField("Original 3", ShapeId: 4)]),
        };
        var revisions = new[]
        {
            new DeckSlideRevision(2, [new DeckField("Revised 2", ShapeId: 3)]),
        };

        var result = JobService.ApplyDeckRevisions(slides, revisions, 50);

        Assert.Equal("Original 1", result[0].Fields[0].Text);
        Assert.Equal("layout-2", result[1].LayoutId);
        Assert.Equal("Revised 2", result[1].Fields[0].Text);
        Assert.Equal("Original 3", result[2].Fields[0].Text);
    }

    [Fact]
    public void RejectsDuplicateDeckRevisionSlideNumbers()
    {
        var slides = new[]
        {
            new DeckSlideSpec("layout-1", [new DeckField("Original", ShapeId: 2)]),
        };
        var revisions = new[]
        {
            new DeckSlideRevision(1, [new DeckField("First", ShapeId: 2)]),
            new DeckSlideRevision(1, [new DeckField("Second", ShapeId: 2)]),
        };

        var error = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyDeckRevisions(slides, revisions, 50));

        Assert.Equal("deck_revision_invalid", error.Code);
    }
}
