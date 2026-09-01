using PptxMcp.Presentation;

namespace PptxMcp.Tests;

public sealed class PptxGenJsVisualPresentationEngineTests
{
    [Theory]
    [InlineData("Failed to launch headless browser: Timed out after 30000 ms while waiting for the WS endpoint URL to appear in stdout!", true)]
    [InlineData("Failed to launch headless browser: missing shared library", false)]
    [InlineData("Model-authored HTML/CSS is invalid", false)]
    public void DetectsOnlyTransientChromiumLaunchTimeouts(string diagnostic, bool expected)
    {
        Assert.Equal(expected, PptxGenJsVisualPresentationEngine.IsTransientBrowserLaunchFailure(diagnostic));
    }
}
