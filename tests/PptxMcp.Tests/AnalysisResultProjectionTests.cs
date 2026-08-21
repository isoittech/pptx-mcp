using System.Text.Json;
using PptxMcp.Domain;
using PptxMcp.Jobs;

namespace PptxMcp.Tests;

public sealed class AnalysisResultProjectionTests
{
    [Fact]
    public void OrdinaryEditingOmitsTheTemplateLayoutCatalog()
    {
        var layout = new LayoutSummary(
            "/ppt/slideLayouts/slideLayout1.xml",
            "Title",
            1,
            [new PlaceholderSummary(2, "Title 1", "title", null)]);
        var summary = new PresentationSummary(
            1,
            false,
            false,
            false,
            false,
            [new SlideSummary(
                1,
                "Company Overview",
                "Title",
                [
                    new ShapeSummary(1, 2, "Title 1", "shape", "Company Overview"),
                    new ShapeSummary(
                        1,
                        4,
                        "Data Table",
                        "table",
                        "Academic MajorsMathematics",
                        ["Academic Majors", "Mathematics"]),
                    new ShapeSummary(1, 3, "Logo", "shape", null),
                ])],
            [layout],
            null,
            []);

        var result = Assert.IsType<TextEditAnalysisResult>(
            JobWorker.PrepareAnalysisResult(summary, includeLayouts: false));

        Assert.False(result.AnalysisTruncated);
        Assert.False(result.HasEditableCharts);
        var slide = Assert.Single(result.Slides);
        Assert.Equal(1, slide[0]);
        Assert.Equal(
            ["Company Overview", "Academic Majors", "Mathematics"],
            Assert.IsType<string[]>(slide[1]));
        var json = JsonSerializer.Serialize(result);
        Assert.DoesNotContain("shape_id", json, StringComparison.Ordinal);
        Assert.DoesNotContain("theme", json, StringComparison.Ordinal);
        Assert.DoesNotContain("kind", json, StringComparison.Ordinal);
        Assert.DoesNotContain("analysis_scope", json, StringComparison.Ordinal);
        Assert.DoesNotContain("entry_schema", json, StringComparison.Ordinal);
        Assert.DoesNotContain("slide_count", json, StringComparison.Ordinal);
        Assert.Contains("\"charts\":false", json, StringComparison.Ordinal);
        Assert.DoesNotContain("validation_errors", json, StringComparison.Ordinal);
    }

    [Fact]
    public void StrictPlaceholderEditingKeepsTheTemplateLayoutCatalog()
    {
        var layout = new LayoutSummary(
            "/ppt/slideLayouts/slideLayout1.xml",
            "Title",
            1,
            [new PlaceholderSummary(2, "Title 1", "title", null)]);
        var summary = new PresentationSummary(
            1,
            false,
            false,
            false,
            false,
            [],
            [layout],
            null,
            []);

        var result = Assert.IsType<PresentationSummary>(
            JobWorker.PrepareAnalysisResult(summary, includeLayouts: true));

        Assert.Same(summary, result);
        Assert.Same(layout, Assert.Single(result.Layouts));
    }
}
