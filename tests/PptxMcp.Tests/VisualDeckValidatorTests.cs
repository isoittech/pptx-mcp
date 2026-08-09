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

    [Fact]
    public void AcceptsDetailedStructuredBriefAndEditableScorecard()
    {
        var deck = new VisualDeckSpec(
            "データセンター投資判断",
            [
                new VisualSlideSpec(
                    VisualSlideKind.StructuredBrief,
                    "需要増と電力制約を同時に解く",
                    Sections:
                    [
                        new VisualBriefSectionSpec(
                            "市場環境",
                            "生成AI需要により高密度ラックへの移行が進む一方、受電容量が拡張速度を制約している。",
                            ["推論需要が継続", "系統接続に長期化リスク"]),
                        new VisualBriefSectionSpec(
                            "経営論点",
                            "設備効率だけでなく、電源調達、立地、冷却方式を一体で設計する必要がある。",
                            Highlight: "統合設計が必要",
                            Tone: "accent"),
                        new VisualBriefSectionSpec(
                            "推奨アクション",
                            Bullets: ["候補地を電力制約で再評価", "液冷対応を標準化", "段階投資で需要変動を吸収"],
                            Tone: "positive"),
                    ],
                    Takeaway: "電力を設備条件ではなく事業ポートフォリオの制約として管理する"),
                new VisualSlideSpec(
                    VisualSlideKind.Scorecard,
                    "候補地を4つの軸で比較",
                    Scorecard: new VisualScorecardSpec(
                        [
                            new VisualScorecardOptionSpec("既存拠点増床", "短期案"),
                            new VisualScorecardOptionSpec("郊外新設", "中期案"),
                            new VisualScorecardOptionSpec("共同利用", "柔軟案"),
                        ],
                        [
                            new VisualScorecardRowSpec("立上げ速度", [
                                new VisualScorecardCellSpec("良", "既存設備を流用", "positive"),
                                new VisualScorecardCellSpec("弱", "許認可を含む", "critical"),
                                new VisualScorecardCellSpec("最良", "契約後に利用可能", "positive"),
                            ]),
                            new VisualScorecardRowSpec("電力余力", [
                                new VisualScorecardCellSpec("条件付", "追加受電が必要", "warning"),
                                new VisualScorecardCellSpec("良", "候補地で確保", "positive"),
                                new VisualScorecardCellSpec("条件付", "事業者に依存", "warning"),
                            ]),
                        ])),
            ],
            Design: new VisualDesignSpec("technical", "detailed", "none"));

        VisualDeckValidator.Validate(deck, 50);
    }

    [Fact]
    public void RejectsStructuredBriefThatExceedsItsReadableTextBudget()
    {
        var deck = new VisualDeckSpec(
            "詳細説明",
            [
                new VisualSlideSpec(
                    VisualSlideKind.StructuredBrief,
                    "一枚で読める量を守る",
                    Sections:
                    [
                        new VisualBriefSectionSpec("論点A", new string('あ', 320)),
                        new VisualBriefSectionSpec("論点B", new string('い', 320)),
                        new VisualBriefSectionSpec("論点C", new string('う', 320)),
                    ]),
            ],
            Design: new VisualDesignSpec(Density: "detailed"));

        var error = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_content_density_invalid", error.Code);
        Assert.Contains("900 total characters", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsScorecardRowsThatDoNotMatchTheOptionCount()
    {
        var deck = new VisualDeckSpec(
            "比較",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Scorecard,
                    "候補比較",
                    Scorecard: new VisualScorecardSpec(
                        [new VisualScorecardOptionSpec("案A"), new VisualScorecardOptionSpec("案B")],
                        [
                            new VisualScorecardRowSpec(
                                "コスト",
                                [new VisualScorecardCellSpec("良")]),
                            new VisualScorecardRowSpec(
                                "速度",
                                [new VisualScorecardCellSpec("良"), new VisualScorecardCellSpec("弱")]),
                        ])),
            ]);

        var error = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_scorecard_cells_mismatch", error.Code);
    }

    [Fact]
    public void WarnsWhenTextRichContentDoesNotUseDetailedDensity()
    {
        var deck = new VisualDeckSpec(
            "説明資料",
            [
                new VisualSlideSpec(
                    VisualSlideKind.StructuredBrief,
                    "情報量に合わせて構造を変える",
                    Sections:
                    [
                        new VisualBriefSectionSpec("論点A", new string('あ', 240)),
                        new VisualBriefSectionSpec("論点B", new string('い', 240)),
                        new VisualBriefSectionSpec("論点C", new string('う', 120)),
                    ]),
            ]);

        VisualDeckValidator.Validate(deck, 50);
        var warnings = VisualDeckValidator.GetDesignWarnings(deck);

        Assert.Contains(warnings, warning => warning.Contains("design.density=detailed", StringComparison.Ordinal));
    }

    [Fact]
    public void AcceptsEditableUkuleleMusicScoreWithFingerColors()
    {
        var deck = new VisualDeckSpec(
            "ウクレレ入門",
            [
                new VisualSlideSpec(
                    VisualSlideKind.MusicScore,
                    "メロディーとTABを対応させる",
                    MusicScore: new VisualMusicScoreSpec(
                        [
                            new VisualMusicMeasureSpec(
                                [
                                    new VisualMusicEventSpec(
                                        "quarter",
                                        [new VisualMusicNoteSpec("C4", 3, 0, 0)]),
                                    new VisualMusicEventSpec(
                                        "eighth",
                                        [new VisualMusicNoteSpec("D4", 3, 2, 2)],
                                        Dotted: true),
                                    new VisualMusicEventSpec(
                                        "eighth",
                                        [new VisualMusicNoteSpec("E4", 2, 0, 0)]),
                                    new VisualMusicEventSpec("quarter", Rest: true),
                                ]),
                            new VisualMusicMeasureSpec(
                                [
                                    new VisualMusicEventSpec(
                                        "half",
                                        [
                                            new VisualMusicNoteSpec("C4", 3, 0, 0),
                                            new VisualMusicNoteSpec("E4", 2, 0, 0),
                                            new VisualMusicNoteSpec("G4", 4, 0, 0),
                                            new VisualMusicNoteSpec("C5", 1, 3, 3),
                                        ]),
                                    new VisualMusicEventSpec(
                                        "quarter",
                                        [new VisualMusicNoteSpec("F4", 2, 1, 1, "#E5484D")]),
                                ],
                                Number: 8),
                        ],
                        TempoBpm: 84,
                        Caption: "色は左手の指番号を示します。")),
            ]);

        VisualDeckValidator.Validate(deck, 50);
    }

    [Fact]
    public void RejectsMusicScoreWhenPitchDoesNotMatchUkuleleTab()
    {
        var deck = new VisualDeckSpec(
            "ウクレレ入門",
            [
                new VisualSlideSpec(
                    VisualSlideKind.MusicScore,
                    "音とフレットを照合する",
                    MusicScore: new VisualMusicScoreSpec(
                        [
                            new VisualMusicMeasureSpec(
                                [
                                    new VisualMusicEventSpec(
                                        "quarter",
                                        [new VisualMusicNoteSpec("D4", 3, 0, 0)]),
                                ]),
                        ])),
            ]);

        var error = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_music_pitch_tab_mismatch", error.Code);
    }

    [Fact]
    public void RejectsMusicScoreWithoutStandardNotationOrTablature()
    {
        var deck = new VisualDeckSpec(
            "ウクレレ入門",
            [
                new VisualSlideSpec(
                    VisualSlideKind.MusicScore,
                    "表示対象を必ず選ぶ",
                    MusicScore: new VisualMusicScoreSpec(
                        [
                            new VisualMusicMeasureSpec(
                                [
                                    new VisualMusicEventSpec(
                                        "quarter",
                                        [new VisualMusicNoteSpec("C4", 3, 0, 0)]),
                                ]),
                        ],
                        ShowStandardNotation: false,
                        ShowTablature: false)),
            ]);

        var error = Assert.Throws<PptxValidationException>(() => VisualDeckValidator.Validate(deck, 50));

        Assert.Equal("visual_music_display_invalid", error.Code);
    }
}
