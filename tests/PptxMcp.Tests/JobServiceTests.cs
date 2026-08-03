using PptxMcp.Jobs;

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
}
