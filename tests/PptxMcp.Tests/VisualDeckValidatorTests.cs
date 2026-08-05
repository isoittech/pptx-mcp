using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class VisualDeckValidatorTests
{
    [Fact]
    public void AppliesCorporateThemeWithoutDiscardingCreativeDirection()
    {
        var deck = new VisualDeckSpec(
            "事業計画",
            [new VisualSlideSpec(VisualSlideKind.Title, "成長戦略")],
            new VisualThemeSpec("aurora", AccentColor: "#ABCDEF"),
            Design: new VisualDesignSpec("editorial", "spacious", "organic"));
        var templateTheme = new PresentationThemeSummary(
            "#112233",
            "#445566",
            "#778899",
            "#FFFFFF",
            "#101010",
            "Corporate Heading",
            "Corporate Body",
            []);

        var result = VisualDeckBranding.ApplyTemplateTheme(deck, templateTheme);

        Assert.Equal("aurora", result.Theme!.Preset);
        Assert.Equal("#112233", result.Theme.PrimaryColor);
        Assert.Equal("#445566", result.Theme.SecondaryColor);
        Assert.Equal("#778899", result.Theme.AccentColor);
        Assert.Equal("#FFFFFF", result.Theme.BackgroundColor);
        Assert.Equal("#101010", result.Theme.TextColor);
        Assert.Equal("Corporate Body", result.Theme.FontFace);
        Assert.Equal(deck.Design, result.Design);
    }

    [Fact]
    public void LeavesVisualDeckUnchangedWhenTemplateHasNoTheme()
    {
        var deck = new VisualDeckSpec(
            "事業計画",
            [new VisualSlideSpec(VisualSlideKind.Title, "成長戦略")]);

        Assert.Same(deck, VisualDeckBranding.ApplyTemplateTheme(deck, null));
    }

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

    [Fact]
    public void AcceptsInfographicLayoutsCreativeDirectionAndVariants()
    {
        var deck = new VisualDeckSpec(
            "サイバー危機対応",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Cards,
                    "経営判断の論点",
                    Cards:
                    [
                        new VisualCardSpec("事業継続", "停止許容時間を決める", "4時間", Icon: "clock"),
                        new VisualCardSpec("封じ込め", "影響範囲を限定する", Icon: "shield"),
                        new VisualCardSpec("対外説明", "一貫した情報を発信する", Icon: "people"),
                    ],
                    Variant: "spotlight"),
                new VisualSlideSpec(
                    VisualSlideKind.Matrix,
                    "対応優先度",
                    Matrix: new VisualMatrixSpec(
                        "実行難易度",
                        "事業インパクト",
                        [
                            new VisualPanelSpec("最優先", ["認証情報の失効"]),
                            new VisualPanelSpec("計画実行", ["代替環境への切替"]),
                            new VisualPanelSpec("監視", ["ログの保全"]),
                            new VisualPanelSpec("後続対応", ["恒久対策"]),
                        ])),
                new VisualSlideSpec(
                    VisualSlideKind.Dashboard,
                    "復旧状況",
                    Metrics:
                    [
                        new VisualMetricSpec("82%", "重要業務の復旧率"),
                        new VisualMetricSpec("14", "残課題", Tone: "warning"),
                    ],
                    Chart: new VisualChartSpec(
                        VisualChartKind.Line,
                        ["0h", "12h", "24h", "48h"],
                        [new VisualChartSeriesSpec("復旧率", [0, 25, 58, 82])])),
            ],
            new VisualThemeSpec("cyber"),
            Design: new VisualDesignSpec("technical", "balanced", "nodes"));

        VisualDeckValidator.Validate(deck, 50);
    }

    [Fact]
    public void AcceptsExpressiveSemanticTonesCustomColorsAndBusinessIcons()
    {
        var deck = new VisualDeckSpec(
            "危機対応",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Dashboard,
                    "経営ダッシュボード",
                    Metrics:
                    [
                        new VisualMetricSpec("3", "重大リスク", Tone: "negative"),
                        new VisualMetricSpec("82%", "復旧率", Tone: "success"),
                        new VisualMetricSpec("12h", "残り時間", Tone: "#7C3AED"),
                    ],
                    Chart: new VisualChartSpec(
                        VisualChartKind.Line,
                        ["0h", "12h"],
                        [new VisualChartSeriesSpec("復旧率", [0, 82])])),
                new VisualSlideSpec(
                    VisualSlideKind.Cards,
                    "判断材料",
                    Cards:
                    [
                        new VisualCardSpec("調査", Icon: "search"),
                        new VisualCardSpec("法令", Tone: "critical", Icon: "compliance"),
                        new VisualCardSpec("意思決定", Icon: "decision"),
                    ]),
            ]);

        VisualDeckValidator.Validate(deck, 50);
    }

    [Fact]
    public void RejectsAnUnknownToneThatCannotBeRenderedIntentionally()
    {
        var deck = new VisualDeckSpec(
            "危機対応",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Metrics,
                    "指標",
                    Metrics:
                    [
                        new VisualMetricSpec("1", "項目A", Tone: "ultraviolet"),
                        new VisualMetricSpec("2", "項目B"),
                    ]),
            ]);

        var error = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_metric_tone_invalid", error.Code);
    }

    [Fact]
    public void AcceptsSixMetricsForAStandaloneMetricsGrid()
    {
        var deck = new VisualDeckSpec(
            "リスク指標",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Metrics,
                    "兆候ダッシュボード",
                    Metrics:
                    [
                        new VisualMetricSpec("1", "項目A"),
                        new VisualMetricSpec("2", "項目B"),
                        new VisualMetricSpec("3", "項目C"),
                        new VisualMetricSpec("4", "項目D"),
                        new VisualMetricSpec("5", "項目E"),
                        new VisualMetricSpec("6", "項目F"),
                    ]),
            ]);

        VisualDeckValidator.Validate(deck, 50);
    }

    [Fact]
    public void KeepsDashboardMetricsAtFourBecauseTheChartSharesTheCanvas()
    {
        var deck = new VisualDeckSpec(
            "KPI",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Dashboard,
                    "KPIダッシュボード",
                    Metrics:
                    [
                        new VisualMetricSpec("1", "項目A"),
                        new VisualMetricSpec("2", "項目B"),
                        new VisualMetricSpec("3", "項目C"),
                        new VisualMetricSpec("4", "項目D"),
                        new VisualMetricSpec("5", "項目E"),
                    ],
                    Chart: new VisualChartSpec(
                        VisualChartKind.Line,
                        ["A", "B"],
                        [new VisualChartSeriesSpec("系列", [1, 2])])),
            ]);

        var error = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_content_density_invalid", error.Code);
        Assert.Contains("more than 4", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsMatrixWithoutFourQuadrants()
    {
        var deck = new VisualDeckSpec(
            "優先度",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Matrix,
                    "対応優先度",
                    Matrix: new VisualMatrixSpec(
                        "難易度",
                        "効果",
                        [
                            new VisualPanelSpec("A", ["項目"]),
                            new VisualPanelSpec("B", ["項目"]),
                            new VisualPanelSpec("C", ["項目"]),
                        ])),
            ]);

        var error = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_matrix_invalid", error.Code);
    }

    [Fact]
    public void ReportsLayoutMonotonyWithoutBlockingGeneration()
    {
        var deck = new VisualDeckSpec(
            "説明資料",
            Enumerable.Range(1, 6)
                .Select(index => new VisualSlideSpec(
                    VisualSlideKind.Bullets,
                    $"論点{index}",
                    Bullets: ["項目A", "項目B"]))
                .ToArray());

        VisualDeckValidator.Validate(deck, 50);
        var warnings = VisualDeckValidator.GetDesignWarnings(deck);

        Assert.Contains(warnings, warning => warning.Contains("four different layout", StringComparison.Ordinal));
        Assert.Contains(warnings, warning => warning.Contains("text-led", StringComparison.Ordinal));
        Assert.Contains(warnings, warning => warning.Contains("repeat the same layout", StringComparison.Ordinal));
    }
}
