using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class VisualDeckValidatorTests
{
    [Fact]
    public void AcceptsSemanticLayoutsAndEditableChartData()
    {
        var deck = new VisualDeckSpec(
            "事業計画",
            [
                new VisualSlideSpec(VisualSlideKind.Title, "成長戦略", Subtitle: "2026年度"),
                new VisualSlideSpec(
                    VisualSlideKind.Metrics,
                    "主要指標",
                    Metrics:
                    [
                        new VisualMetricSpec("+18%", "売上成長率"),
                        new VisualMetricSpec("4.8", "顧客満足度", "5点満点", "positive"),
                    ]),
                new VisualSlideSpec(
                    VisualSlideKind.Chart,
                    "四半期別売上",
                    Chart: new VisualChartSpec(
                        VisualChartKind.Bar,
                        ["Q1", "Q2", "Q3", "Q4"],
                        [new VisualChartSeriesSpec("売上", [10, 14, 17, 22])])),
            ],
            new VisualThemeSpec("aurora"));

        VisualDeckValidator.Validate(deck, 50);
    }

    [Fact]
    public void RejectsAChartWhoseSeriesDoesNotMatchCategories()
    {
        var deck = new VisualDeckSpec(
            "事業計画",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Chart,
                    "四半期別売上",
                    Chart: new VisualChartSpec(
                        VisualChartKind.Line,
                        ["Q1", "Q2", "Q3"],
                        [new VisualChartSeriesSpec("売上", [10, 14])])),
            ]);

        var exception = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_chart_values_mismatch", exception.Code);
    }

    [Fact]
    public void RejectsContentThatWouldMakeAProcessSlideTooDense()
    {
        var deck = new VisualDeckSpec(
            "移行計画",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Process,
                    "導入工程",
                    Steps: Enumerable.Range(1, 7)
                        .Select(index => new VisualStepSpec($"工程{index}"))
                        .ToArray()),
            ]);

        var exception = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_content_density_invalid", exception.Code);
    }

    [Fact]
    public void RejectsInvalidCustomThemeColor()
    {
        var deck = new VisualDeckSpec(
            "事業計画",
            [new VisualSlideSpec(VisualSlideKind.Title, "成長戦略")],
            new VisualThemeSpec(PrimaryColor: "navy"));

        var exception = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_color_invalid", exception.Code);
    }

    [Fact]
    public void AcceptsCustomThemeColorsWithLeadingHash()
    {
        var deck = new VisualDeckSpec(
            "事業計画",
            [new VisualSlideSpec(VisualSlideKind.Title, "成長戦略")],
            new VisualThemeSpec(
                PrimaryColor: "#17213A",
                AccentColor: "#67E8F9"));

        VisualDeckValidator.Validate(deck, 50);
    }
}
