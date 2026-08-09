using System.ComponentModel;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using PptxMcp.Storage;

namespace PptxMcp.Domain;

[JsonConverter(typeof(JsonStringEnumConverter<VisualSlideKind>))]
public enum VisualSlideKind
{
    Title,
    Agenda,
    Section,
    Bullets,
    Metrics,
    Comparison,
    Process,
    Timeline,
    Chart,
    Statement,
    Cards,
    Matrix,
    Funnel,
    Roadmap,
    Dashboard,
    Quote,
    Closing,
    StructuredBrief,
    Scorecard,
    MusicScore,
}

[JsonConverter(typeof(JsonStringEnumConverter<VisualChartKind>))]
public enum VisualChartKind
{
    Bar,
    Line,
    Pie,
    Doughnut,
}

public sealed record VisualDeckSpec(
    string Title,
    IReadOnlyList<VisualSlideSpec> Slides,
    VisualThemeSpec? Theme = null,
    string? Subject = null,
    string Language = "ja-JP",
    VisualDesignSpec? Design = null);

public sealed record VisualDeckDraftView(
    [property: JsonPropertyName("draft_id")] string DraftId,
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("expected_slide_count")] int ExpectedSlideCount,
    [property: JsonPropertyName("accepted_slide_count")] int AcceptedSlideCount,
    [property: JsonPropertyName("next_slide_number")] int NextSlideNumber,
    [property: JsonPropertyName("remaining_slide_count")] int RemainingSlideCount,
    [property: JsonPropertyName("maximum_batch_slides")] int MaximumBatchSlides,
    [property: JsonPropertyName("instruction")] string Instruction);

public sealed record BrandedVisualDeckSpec(
    [property: JsonPropertyName("deck")] VisualDeckSpec Deck,
    [property: JsonPropertyName("template_layout_id")] string TemplateLayoutId = "auto");

public sealed record VisualDesignSpec(
    string Style = "executive",
    string Density = "balanced",
    string Motif = "geometric");

public sealed record VisualThemeSpec(
    string Preset = "midnight",
    string? PrimaryColor = null,
    string? SecondaryColor = null,
    string? AccentColor = null,
    string? BackgroundColor = null,
    string? TextColor = null,
    string? FontFace = null);

public sealed record VisualSlideSpec(
    VisualSlideKind Kind,
    string Title,
    string? Subtitle = null,
    string? Eyebrow = null,
    string? Body = null,
    IReadOnlyList<string>? Bullets = null,
    IReadOnlyList<VisualMetricSpec>? Metrics = null,
    IReadOnlyList<VisualPanelSpec>? Panels = null,
    IReadOnlyList<VisualStepSpec>? Steps = null,
    VisualChartSpec? Chart = null,
    string? Attribution = null,
    string? Takeaway = null,
    IReadOnlyList<VisualCardSpec>? Cards = null,
    VisualMatrixSpec? Matrix = null,
    [property: Description("Two or three titled content columns for a text-rich but scannable structured_brief slide. Use sections instead of one long body.")]
    IReadOnlyList<VisualBriefSectionSpec>? Sections = null,
    [property: Description("Editable comparison table for a scorecard slide. Each criterion must have exactly one cell per option.")]
    VisualScorecardSpec? Scorecard = null,
    string Variant = "auto",
    [property: Description("Editable native standard-notation and ukulele TAB score for a music_score slide. Use semantic pitches, durations, strings, frets, and fingers instead of coordinates.")]
    VisualMusicScoreSpec? MusicScore = null);

public sealed record VisualMusicScoreSpec(
    [property: Description("One to eight measures in reading order. A slide may contain at most 64 events in total.")]
    IReadOnlyList<VisualMusicMeasureSpec> Measures,
    [property: Description("Time signature such as 4/4, 3/4, or 6/8.")]
    string TimeSignature = "4/4",
    [property: Description("Short display label for the key, such as C, G, Am, or F#m.")]
    string KeySignature = "C",
    [property: Description("Only treble is supported in the first editable music-score layout.")]
    string Clef = "treble",
    [property: Description("Ukulele tuning: high-g or low-g. String numbers are 1=A, 2=E, 3=C, and 4=G.")]
    string Tuning = "high-g",
    [property: Description("Optional tempo in beats per minute, from 20 to 300.")]
    int? TempoBpm = null,
    [property: Description("Draw the five-line standard staff with editable noteheads, stems, rests, accidentals, and ledger lines.")]
    bool ShowStandardNotation = true,
    [property: Description("Draw the four-line ukulele TAB staff with editable fret-number text.")]
    bool ShowTablature = true,
    [property: Description("Color TAB fret markers by left-hand finger: 0=open, 1=index, 2=middle, 3=ring, 4=little.")]
    bool ColorFingerings = true,
    [property: Description("Optional concise explanation shown below the score.")]
    string? Caption = null);

public sealed record VisualMusicMeasureSpec(
    [property: Description("One to twelve sequential musical events in this measure.")]
    IReadOnlyList<VisualMusicEventSpec> Events,
    [property: Description("Optional displayed measure number. When omitted, measures are numbered from one.")]
    int? Number = null);

public sealed record VisualMusicEventSpec(
    [property: Description("Duration: whole, half, quarter, eighth, or sixteenth.")]
    string Duration,
    [property: Description("One to four simultaneous notes. Omit notes only when rest=true.")]
    IReadOnlyList<VisualMusicNoteSpec>? Notes = null,
    [property: Description("Draw an editable rest symbol instead of notes.")]
    bool Rest = false,
    [property: Description("Add an augmentation dot to the event.")]
    bool Dotted = false,
    [property: Description("Draw a tie toward the next event in the same measure.")]
    bool TieToNext = false,
    [property: Description("Optional short performance annotation above the event.")]
    string? Annotation = null);

public sealed record VisualMusicNoteSpec(
    [property: Description("Scientific pitch notation from G3 through C6, for example C4, F#4, or Bb4.")]
    string Pitch,
    [property: JsonPropertyName("string"), Description("Ukulele string number: 1=A, 2=E, 3=C, 4=G.")]
    int StringNumber,
    [property: Description("Fret number from 0 through 24.")]
    int Fret,
    [property: Description("Left-hand finger: 0=open, 1=index, 2=middle, 3=ring, 4=little.")]
    int? Finger = null,
    [property: Description("Optional #RRGGBB override for this note's TAB marker.")]
    string? Color = null);

public sealed record VisualBriefSectionSpec(
    [property: Description("Short heading that lets readers understand this block while scanning headings only.")]
    string Heading,
    [property: Description("Concise explanatory prose for this block. Keep one main point per section.")]
    string? Body = null,
    [property: Description("Optional native bullet list when parallel facts are clearer than prose.")]
    IReadOnlyList<string>? Bullets = null,
    [property: Description("Optional short label for the single most important fact in this section.")]
    string? Highlight = null,
    [property: Description("Semantic tone for the section rule or highlight. Prefer neutral for most sections and emphasize only exceptions.")]
    string Tone = "neutral");

public sealed record VisualScorecardSpec(
    [property: Description("Two to four options shown as editable table columns, in reading order.")]
    IReadOnlyList<VisualScorecardOptionSpec> Options,
    [property: Description("Two to six evaluation criteria shown as editable table rows.")]
    IReadOnlyList<VisualScorecardRowSpec> Criteria);

public sealed record VisualScorecardOptionSpec(
    string Title,
    string? Subtitle = null);

public sealed record VisualScorecardRowSpec(
    string Criterion,
    [property: Description("One assessment cell per option, in exactly the same order as scorecard.options.")]
    IReadOnlyList<VisualScorecardCellSpec> Cells);

public sealed record VisualScorecardCellSpec(
    [property: Description("Short assessment such as Recommended, Good, Conditional, or Poor.")]
    string Rating,
    [property: Description("Short evidence or rationale supporting the rating.")]
    string? Detail = null,
    [property: Description("Semantic tone such as positive, warning, critical, neutral, or a custom #RRGGBB color.")]
    string Tone = "neutral");

public sealed record VisualMetricSpec(
    string Value,
    string Label,
    string? Detail = null,
    [property: Description("Semantic tone such as accent, positive/success, warning, critical/danger/negative, neutral/muted, info, or a custom #RRGGBB color.")]
    string Tone = "accent");

public sealed record VisualPanelSpec(
    string Title,
    IReadOnlyList<string> Bullets,
    string? Highlight = null);

public sealed record VisualStepSpec(
    string Title,
    string? Description = null,
    string? Label = null);

public sealed record VisualCardSpec(
    string Title,
    string? Description = null,
    string? Value = null,
    [property: Description("Semantic tone such as accent, positive/success, warning, critical/danger/negative, neutral/muted, info, or a custom #RRGGBB color.")]
    string Tone = "accent",
    [property: Description("Editable native icon: insight, target, growth, people, shield, clock, cloud, settings, data, warning, check, idea, search, compliance, decision, lock, network, document, communication, recovery, backup, legal, monitor, or automation.")]
    string Icon = "insight");

public sealed record VisualMatrixSpec(
    string HorizontalAxis,
    string VerticalAxis,
    IReadOnlyList<VisualPanelSpec> Quadrants);

public sealed record VisualChartSpec(
    VisualChartKind Kind,
    IReadOnlyList<string> Categories,
    IReadOnlyList<VisualChartSeriesSpec> Series,
    string? ValueSuffix = null,
    bool ShowLegend = true);

public sealed record VisualChartSeriesSpec(
    string Name,
    IReadOnlyList<double> Values);

public static class VisualDeckBranding
{
    public static VisualDeckSpec ApplyTemplateTheme(
        VisualDeckSpec deck,
        PresentationThemeSummary? templateTheme)
    {
        if (templateTheme is null)
        {
            return deck;
        }

        var current = deck.Theme ?? new VisualThemeSpec("minimal");
        var brandedTheme = current with
        {
            PrimaryColor = templateTheme.PrimaryColor ?? current.PrimaryColor,
            SecondaryColor = templateTheme.SecondaryColor ?? current.SecondaryColor,
            AccentColor = templateTheme.AccentColor ?? current.AccentColor,
            BackgroundColor = templateTheme.BackgroundColor ?? current.BackgroundColor,
            TextColor = templateTheme.TextColor ?? current.TextColor,
            FontFace = templateTheme.BodyFont ?? templateTheme.HeadingFont ?? current.FontFace,
        };
        return deck with { Theme = brandedTheme };
    }
}

public sealed record VisualDeckCreationResult(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("layout_kinds")] IReadOnlyList<string> LayoutKinds,
    [property: JsonPropertyName("renderer")] string Renderer,
    [property: JsonPropertyName("design_warnings")] IReadOnlyList<string> DesignWarnings);

public sealed record BrandedVisualDeckCreationResult(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("layout_kinds")] IReadOnlyList<string> LayoutKinds,
    [property: JsonPropertyName("renderer")] string Renderer,
    [property: JsonPropertyName("template_layout_id")] string TemplateLayoutId,
    [property: JsonPropertyName("template_layout_name")] string TemplateLayoutName,
    [property: JsonPropertyName("template_theme_applied")] bool TemplateThemeApplied,
    [property: JsonPropertyName("design_warnings")] IReadOnlyList<string> DesignWarnings);

public sealed record VisualSlideRevision(
    [property: JsonPropertyName("slide_number"), Description("One-based slide number to replace.")] int SlideNumber,
    [property: JsonPropertyName("slide"), Description("Complete replacement VisualSlideSpec for this one slide.")] VisualSlideSpec Slide);

public static partial class VisualDeckValidator
{
    private static readonly HashSet<string> Presets = new(StringComparer.OrdinalIgnoreCase)
    {
        "midnight",
        "aurora",
        "sunset",
        "forest",
        "minimal",
        "ocean",
        "berry",
        "clay",
        "cyber",
    };

    private static readonly HashSet<string> SemanticTones = new(StringComparer.OrdinalIgnoreCase)
    {
        "accent",
        "primary",
        "secondary",
        "info",
        "positive",
        "success",
        "warning",
        "critical",
        "danger",
        "negative",
        "risk",
        "neutral",
        "muted",
    };

    private static readonly HashSet<string> DesignStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        "executive",
        "editorial",
        "bold",
        "technical",
        "playful",
    };

    private static readonly HashSet<string> DesignDensities = new(StringComparer.OrdinalIgnoreCase)
    {
        "airy",
        "balanced",
        "detailed",
    };

    private static readonly HashSet<string> DesignMotifs = new(StringComparer.OrdinalIgnoreCase)
    {
        "geometric",
        "orbit",
        "nodes",
        "ribbon",
        "none",
    };

    private static readonly HashSet<string> CardIcons = new(StringComparer.OrdinalIgnoreCase)
    {
        "insight",
        "target",
        "growth",
        "people",
        "shield",
        "clock",
        "cloud",
        "settings",
        "data",
        "warning",
        "check",
        "idea",
        "search",
        "compliance",
        "decision",
        "lock",
        "network",
        "document",
        "communication",
        "recovery",
        "backup",
        "legal",
        "monitor",
        "automation",
    };

    private static readonly HashSet<string> SlideVariants = new(StringComparer.OrdinalIgnoreCase)
    {
        "auto",
        "grid",
        "spotlight",
        "split",
        "cascade",
        "editorial",
    };

    private static readonly HashSet<string> MusicDurations = new(StringComparer.OrdinalIgnoreCase)
    {
        "whole",
        "half",
        "quarter",
        "eighth",
        "sixteenth",
    };

    private static readonly HashSet<string> UkuleleTunings = new(StringComparer.OrdinalIgnoreCase)
    {
        "high-g",
        "low-g",
    };

    public static void Validate(VisualDeckSpec deck, int maximumSlides)
    {
        ArgumentNullException.ThrowIfNull(deck);
        ValidateMetadata(deck.Title, deck.Subject, deck.Language, deck.Theme, deck.Design);

        if (deck.Slides is null || deck.Slides.Count is < 1 || deck.Slides.Count > maximumSlides)
        {
            throw new PptxValidationException(
                "visual_slide_count_out_of_range",
                $"Visual decks must contain between 1 and {maximumSlides} slides.");
        }

        for (var index = 0; index < deck.Slides.Count; index++)
        {
            ValidateSlide(deck.Slides[index], index + 1);
        }
    }

    public static void ValidateMetadata(
        string title,
        string? subject,
        string language,
        VisualThemeSpec? theme,
        VisualDesignSpec? design)
    {
        ValidateText(title, "deck.title", 1, 160);
        ValidateOptionalText(subject, "deck.subject", 240);
        ValidateText(language, "deck.language", 2, 24);
        ValidateTheme(theme);
        ValidateDesign(design);
    }

    public static IReadOnlyList<string> GetDesignWarnings(VisualDeckSpec deck)
    {
        var warnings = new List<string>();
        if (deck.Slides.Count >= 6)
        {
            var distinctKinds = deck.Slides.Select(static slide => slide.Kind).Distinct().Count();
            if (distinctKinds < 4)
            {
                warnings.Add("Use at least four different layout kinds in decks of six or more slides.");
            }

            var textLedSlides = deck.Slides.Count(static slide =>
                slide.Kind is VisualSlideKind.Bullets or VisualSlideKind.Quote or VisualSlideKind.Statement);
            if (textLedSlides * 2 > deck.Slides.Count)
            {
                warnings.Add("More than half of the deck is text-led; replace some slides with cards, data, process, matrix, or roadmap visuals.");
            }
        }

        for (var index = 2; index < deck.Slides.Count; index++)
        {
            if (deck.Slides[index].Kind == deck.Slides[index - 1].Kind
                && deck.Slides[index].Kind == deck.Slides[index - 2].Kind)
            {
                warnings.Add($"Slides {index - 1}-{index + 1} repeat the same layout kind; vary the visual rhythm.");
            }
        }

        var duplicateTitles = deck.Slides
            .GroupBy(static slide => slide.Title.Trim(), StringComparer.OrdinalIgnoreCase)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .ToArray();
        if (duplicateTitles.Length > 0)
        {
            warnings.Add("Some slide titles are duplicated; make the heading sequence communicate the story without reading body text.");
        }

        var usesDetailedDensity = string.Equals(
            deck.Design?.Density,
            "detailed",
            StringComparison.OrdinalIgnoreCase);
        for (var index = 0; index < deck.Slides.Count; index++)
        {
            var slide = deck.Slides[index];
            if (GetTextCharacterCount(slide) >= 500 && !usesDetailedDensity)
            {
                warnings.Add($"Slide {index + 1} is text-rich; use design.density=detailed and a structured_brief or scorecard layout instead of shrinking text.");
            }

            if (slide.Kind == VisualSlideKind.StructuredBrief
                && slide.Sections is { Count: > 1 }
                && slide.Sections.All(static section =>
                    !string.IsNullOrWhiteSpace(section.Highlight)
                    || IsStrongEmphasisTone(section.Tone)))
            {
                warnings.Add($"Slide {index + 1} emphasizes every section; reserve strong color or highlight labels for the few items that need attention.");
            }
        }

        return warnings;
    }

    private static int GetTextCharacterCount(VisualSlideSpec slide)
    {
        var count = slide.Title.Length
            + (slide.Subtitle?.Length ?? 0)
            + (slide.Eyebrow?.Length ?? 0)
            + (slide.Body?.Length ?? 0)
            + (slide.Attribution?.Length ?? 0)
            + (slide.Takeaway?.Length ?? 0)
            + (slide.Bullets?.Sum(static item => item.Length) ?? 0)
            + (slide.Metrics?.Sum(static item => item.Value.Length + item.Label.Length + (item.Detail?.Length ?? 0)) ?? 0)
            + (slide.Panels?.Sum(static item => item.Title.Length + (item.Highlight?.Length ?? 0) + item.Bullets.Sum(static bullet => bullet.Length)) ?? 0)
            + (slide.Steps?.Sum(static item => item.Title.Length + (item.Description?.Length ?? 0) + (item.Label?.Length ?? 0)) ?? 0)
            + (slide.Cards?.Sum(static item => item.Title.Length + (item.Description?.Length ?? 0) + (item.Value?.Length ?? 0)) ?? 0)
            + (slide.Sections?.Sum(static item => item.Heading.Length + (item.Body?.Length ?? 0) + (item.Highlight?.Length ?? 0) + (item.Bullets?.Sum(static bullet => bullet.Length) ?? 0)) ?? 0)
            + (slide.MusicScore?.Caption?.Length ?? 0)
            + (slide.MusicScore?.Measures.Sum(static measure =>
                measure.Events.Sum(static item => item.Annotation?.Length ?? 0)) ?? 0);

        if (slide.Scorecard is not null)
        {
            count += slide.Scorecard.Options.Sum(static option => option.Title.Length + (option.Subtitle?.Length ?? 0));
            count += slide.Scorecard.Criteria.Sum(static row =>
                row.Criterion.Length
                + row.Cells.Sum(static cell => cell.Rating.Length + (cell.Detail?.Length ?? 0)));
        }

        return count;
    }

    private static bool IsStrongEmphasisTone(string tone) =>
        tone.Equals("positive", StringComparison.OrdinalIgnoreCase)
        || tone.Equals("success", StringComparison.OrdinalIgnoreCase)
        || tone.Equals("warning", StringComparison.OrdinalIgnoreCase)
        || tone.Equals("critical", StringComparison.OrdinalIgnoreCase)
        || tone.Equals("danger", StringComparison.OrdinalIgnoreCase)
        || tone.Equals("negative", StringComparison.OrdinalIgnoreCase)
        || tone.Equals("risk", StringComparison.OrdinalIgnoreCase)
        || HexColorRegex().IsMatch(tone);

    private static void ValidateDesign(VisualDesignSpec? design)
    {
        if (design is null)
        {
            return;
        }

        if (!DesignStyles.Contains(design.Style)
            || !DesignDensities.Contains(design.Density)
            || !DesignMotifs.Contains(design.Motif))
        {
            throw new PptxValidationException(
                "visual_design_invalid",
                "design must use a supported style, density, and motif.");
        }
    }

    private static void ValidateTheme(VisualThemeSpec? theme)
    {
        if (theme is null)
        {
            return;
        }

        if (!Presets.Contains(theme.Preset))
        {
            throw new PptxValidationException(
                "visual_theme_invalid",
                "theme.preset must be one of midnight, aurora, sunset, forest, minimal, ocean, berry, clay, or cyber.");
        }

        ValidateColor(theme.PrimaryColor, "theme.primaryColor");
        ValidateColor(theme.SecondaryColor, "theme.secondaryColor");
        ValidateColor(theme.AccentColor, "theme.accentColor");
        ValidateColor(theme.BackgroundColor, "theme.backgroundColor");
        ValidateColor(theme.TextColor, "theme.textColor");
        ValidateOptionalText(theme.FontFace, "theme.fontFace", 80);
    }

    private static void ValidateSlide(VisualSlideSpec slide, int slideNumber)
    {
        if (slide is null)
        {
            throw new PptxValidationException("visual_slide_invalid", $"Slide {slideNumber} is null.");
        }

        var prefix = $"slides[{slideNumber - 1}]";
        ValidateText(slide.Title, $"{prefix}.title", 1, 140);
        ValidateOptionalText(slide.Subtitle, $"{prefix}.subtitle", 280);
        ValidateOptionalText(slide.Eyebrow, $"{prefix}.eyebrow", 80);
        ValidateOptionalText(slide.Body, $"{prefix}.body", 700);
        ValidateOptionalText(slide.Attribution, $"{prefix}.attribution", 120);
        ValidateOptionalText(slide.Takeaway, $"{prefix}.takeaway", 280);
        if (!SlideVariants.Contains(slide.Variant))
        {
            throw new PptxValidationException(
                "visual_slide_variant_invalid",
                $"{prefix}.variant must be auto, grid, spotlight, split, cascade, or editorial.");
        }

        ValidateList(slide.Bullets, $"{prefix}.bullets", 8, 180);

        if (slide.Metrics is not null)
        {
            var maximumMetrics = slide.Kind == VisualSlideKind.Metrics ? 6 : 4;
            if (slide.Metrics.Count > maximumMetrics)
            {
                ThrowDensity(prefix, "metrics", maximumMetrics);
            }

            foreach (var metric in slide.Metrics)
            {
                ValidateText(metric.Value, $"{prefix}.metrics.value", 1, 32);
                ValidateText(metric.Label, $"{prefix}.metrics.label", 1, 80);
                ValidateOptionalText(metric.Detail, $"{prefix}.metrics.detail", 140);
                if (!IsSupportedTone(metric.Tone))
                {
                    throw new PptxValidationException(
                        "visual_metric_tone_invalid",
                        $"{prefix}.metrics.tone must be a supported semantic tone or a #RRGGBB color.");
                }
            }
        }

        if (slide.Panels is not null)
        {
            if (slide.Panels.Count > 3)
            {
                ThrowDensity(prefix, "panels", 3);
            }

            foreach (var panel in slide.Panels)
            {
                ValidateText(panel.Title, $"{prefix}.panels.title", 1, 80);
                ValidateList(panel.Bullets, $"{prefix}.panels.bullets", 6, 140, minimumCount: 1);
                ValidateOptionalText(panel.Highlight, $"{prefix}.panels.highlight", 80);
            }
        }

        if (slide.Steps is not null)
        {
            if (slide.Steps.Count > 6)
            {
                ThrowDensity(prefix, "steps", 6);
            }

            foreach (var step in slide.Steps)
            {
                ValidateText(step.Title, $"{prefix}.steps.title", 1, 72);
                ValidateOptionalText(step.Description, $"{prefix}.steps.description", 160);
                ValidateOptionalText(step.Label, $"{prefix}.steps.label", 40);
            }
        }

        if (slide.Cards is not null)
        {
            if (slide.Cards.Count > 6)
            {
                ThrowDensity(prefix, "cards", 6);
            }

            foreach (var card in slide.Cards)
            {
                ValidateText(card.Title, $"{prefix}.cards.title", 1, 72);
                ValidateOptionalText(card.Description, $"{prefix}.cards.description", 150);
                ValidateOptionalText(card.Value, $"{prefix}.cards.value", 32);
                if (!IsSupportedTone(card.Tone))
                {
                    throw new PptxValidationException(
                        "visual_card_tone_invalid",
                        $"{prefix}.cards.tone must be a supported semantic tone or a #RRGGBB color.");
                }

                if (!CardIcons.Contains(card.Icon))
                {
                    throw new PptxValidationException(
                        "visual_card_icon_invalid",
                        $"{prefix}.cards.icon is not supported.");
                }
            }
        }

        ValidateBriefSections(slide.Sections, prefix);
        ValidateScorecard(slide.Scorecard, prefix);
        ValidateMatrix(slide.Matrix, prefix);
        ValidateMusicScore(slide.MusicScore, prefix);

        ValidateChart(slide.Chart, prefix);
        switch (slide.Kind)
        {
            case VisualSlideKind.Agenda:
                RequireCount(slide.Bullets, prefix, "bullets", 2, 8);
                break;
            case VisualSlideKind.Bullets:
                RequireCount(slide.Bullets, prefix, "bullets", 2, 7);
                break;
            case VisualSlideKind.Metrics:
                RequireCount(slide.Metrics, prefix, "metrics", 2, 6);
                break;
            case VisualSlideKind.Comparison:
                RequireCount(slide.Panels, prefix, "panels", 2, 3);
                break;
            case VisualSlideKind.Process:
            case VisualSlideKind.Timeline:
            case VisualSlideKind.Funnel:
            case VisualSlideKind.Roadmap:
                RequireCount(slide.Steps, prefix, "steps", 3, 6);
                break;
            case VisualSlideKind.Chart when slide.Chart is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.chart is required for a chart slide.");
            case VisualSlideKind.Statement when string.IsNullOrWhiteSpace(slide.Body):
                throw new PptxValidationException("visual_content_missing", $"{prefix}.body is required for a statement slide.");
            case VisualSlideKind.Cards:
                RequireCount(slide.Cards, prefix, "cards", 3, 6);
                break;
            case VisualSlideKind.Matrix when slide.Matrix is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.matrix is required for a matrix slide.");
            case VisualSlideKind.Dashboard:
                RequireCount(slide.Metrics, prefix, "metrics", 2, 4);
                if (slide.Chart is null)
                {
                    throw new PptxValidationException("visual_content_missing", $"{prefix}.chart is required for a dashboard slide.");
                }

                break;
            case VisualSlideKind.Quote when string.IsNullOrWhiteSpace(slide.Body):
                throw new PptxValidationException("visual_content_missing", $"{prefix}.body is required for a quote slide.");
            case VisualSlideKind.StructuredBrief:
                RequireCount(slide.Sections, prefix, "sections", 2, 3);
                break;
            case VisualSlideKind.Scorecard when slide.Scorecard is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.scorecard is required for a scorecard slide.");
            case VisualSlideKind.MusicScore when slide.MusicScore is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.musicScore is required for a musicScore slide.");
        }
    }

    private static void ValidateMusicScore(VisualMusicScoreSpec? musicScore, string prefix)
    {
        if (musicScore is null)
        {
            return;
        }

        var path = $"{prefix}.musicScore";
        if (musicScore.Measures is null || musicScore.Measures.Count is < 1 or > 8)
        {
            throw new PptxValidationException(
                "visual_music_measures_out_of_range",
                $"{path}.measures must contain between 1 and 8 measures.");
        }

        if (!MusicTimeSignatureRegex().IsMatch(musicScore.TimeSignature))
        {
            throw new PptxValidationException(
                "visual_music_time_signature_invalid",
                $"{path}.timeSignature must use a numerator from 1 to 99 and a denominator of 1, 2, 4, 8, or 16.");
        }

        ValidateText(musicScore.KeySignature, $"{path}.keySignature", 1, 16);
        if (!musicScore.Clef.Equals("treble", StringComparison.OrdinalIgnoreCase))
        {
            throw new PptxValidationException(
                "visual_music_clef_invalid",
                $"{path}.clef must be treble.");
        }

        if (!UkuleleTunings.Contains(musicScore.Tuning))
        {
            throw new PptxValidationException(
                "visual_music_tuning_invalid",
                $"{path}.tuning must be high-g or low-g.");
        }

        if (musicScore.TempoBpm is < 20 or > 300)
        {
            throw new PptxValidationException(
                "visual_music_tempo_invalid",
                $"{path}.tempoBpm must be between 20 and 300.");
        }

        if (!musicScore.ShowStandardNotation && !musicScore.ShowTablature)
        {
            throw new PptxValidationException(
                "visual_music_display_invalid",
                $"{path} must enable standard notation, tablature, or both.");
        }

        ValidateOptionalText(musicScore.Caption, $"{path}.caption", 180);
        var totalEventCount = 0;
        for (var measureIndex = 0; measureIndex < musicScore.Measures.Count; measureIndex++)
        {
            var measure = musicScore.Measures[measureIndex];
            var measurePath = $"{path}.measures[{measureIndex}]";
            if (measure.Events is null || measure.Events.Count is < 1 or > 12)
            {
                throw new PptxValidationException(
                    "visual_music_events_out_of_range",
                    $"{measurePath}.events must contain between 1 and 12 events.");
            }

            if (measure.Number is < 1 or > 999)
            {
                throw new PptxValidationException(
                    "visual_music_measure_number_invalid",
                    $"{measurePath}.number must be between 1 and 999.");
            }

            totalEventCount += measure.Events.Count;
            for (var eventIndex = 0; eventIndex < measure.Events.Count; eventIndex++)
            {
                ValidateMusicEvent(
                    measure.Events[eventIndex],
                    $"{measurePath}.events[{eventIndex}]",
                    musicScore.Tuning,
                    eventIndex == measure.Events.Count - 1);
            }
        }

        if (totalEventCount > 64)
        {
            throw new PptxValidationException(
                "visual_music_event_density_invalid",
                $"{path} must not contain more than 64 events on one slide.");
        }
    }

    private static void ValidateMusicEvent(
        VisualMusicEventSpec musicEvent,
        string path,
        string tuning,
        bool isLastEvent)
    {
        if (!MusicDurations.Contains(musicEvent.Duration))
        {
            throw new PptxValidationException(
                "visual_music_duration_invalid",
                $"{path}.duration must be whole, half, quarter, eighth, or sixteenth.");
        }

        ValidateOptionalText(musicEvent.Annotation, $"{path}.annotation", 32);
        if (musicEvent.TieToNext && isLastEvent)
        {
            throw new PptxValidationException(
                "visual_music_tie_invalid",
                $"{path}.tieToNext requires another event in the same measure.");
        }

        var notes = musicEvent.Notes;
        if (musicEvent.Rest)
        {
            if (notes is { Count: > 0 })
            {
                throw new PptxValidationException(
                    "visual_music_rest_invalid",
                    $"{path}.notes must be omitted or empty when rest=true.");
            }

            return;
        }

        if (notes is null || notes.Count is < 1 or > 4)
        {
            throw new PptxValidationException(
                "visual_music_notes_out_of_range",
                $"{path}.notes must contain between 1 and 4 notes when rest=false.");
        }

        if (notes.Select(static note => note.StringNumber).Distinct().Count() != notes.Count)
        {
            throw new PptxValidationException(
                "visual_music_string_duplicate",
                $"{path}.notes must use each ukulele string at most once per event.");
        }

        for (var noteIndex = 0; noteIndex < notes.Count; noteIndex++)
        {
            ValidateMusicNote(notes[noteIndex], $"{path}.notes[{noteIndex}]", tuning);
        }
    }

    private static void ValidateMusicNote(VisualMusicNoteSpec note, string path, string tuning)
    {
        var match = MusicPitchRegex().Match(note.Pitch);
        if (!match.Success || !TryGetMidiPitch(match, out var midiPitch) || midiPitch is < 55 or > 84)
        {
            throw new PptxValidationException(
                "visual_music_pitch_invalid",
                $"{path}.pitch must use scientific pitch notation from G3 through C6.");
        }

        if (note.StringNumber is < 1 or > 4)
        {
            throw new PptxValidationException(
                "visual_music_string_invalid",
                $"{path}.string must be between 1 and 4.");
        }

        if (note.Fret is < 0 or > 24)
        {
            throw new PptxValidationException(
                "visual_music_fret_invalid",
                $"{path}.fret must be between 0 and 24.");
        }

        if (note.Finger is < 0 or > 4)
        {
            throw new PptxValidationException(
                "visual_music_finger_invalid",
                $"{path}.finger must be between 0 and 4.");
        }

        if (note.Fret == 0 && note.Finger is > 0)
        {
            throw new PptxValidationException(
                "visual_music_finger_invalid",
                $"{path}.finger must be 0 or omitted for an open string.");
        }

        ValidateColor(note.Color, $"{path}.color");
        var openPitch = note.StringNumber switch
        {
            1 => 69,
            2 => 64,
            3 => 60,
            4 when tuning.Equals("low-g", StringComparison.OrdinalIgnoreCase) => 55,
            4 => 67,
            _ => throw new InvalidOperationException("The ukulele string was validated before pitch comparison."),
        };
        if (openPitch + note.Fret != midiPitch)
        {
            throw new PptxValidationException(
                "visual_music_pitch_tab_mismatch",
                $"{path}.pitch does not match its string and fret for {tuning} tuning.");
        }
    }

    private static bool TryGetMidiPitch(Match match, out int midiPitch)
    {
        var semitone = char.ToUpperInvariant(match.Groups[1].Value[0]) switch
        {
            'C' => 0,
            'D' => 2,
            'E' => 4,
            'F' => 5,
            'G' => 7,
            'A' => 9,
            'B' => 11,
            _ => -100,
        };
        var accidental = match.Groups[2].Value switch
        {
            "#" => 1,
            "b" => -1,
            _ => 0,
        };
        if (!int.TryParse(match.Groups[3].Value, out var octave) || semitone < 0)
        {
            midiPitch = 0;
            return false;
        }

        midiPitch = (octave + 1) * 12 + semitone + accidental;
        return true;
    }

    private static void ValidateBriefSections(
        IReadOnlyList<VisualBriefSectionSpec>? sections,
        string prefix)
    {
        if (sections is null)
        {
            return;
        }

        if (sections.Count > 3)
        {
            ThrowDensity(prefix, "sections", 3);
        }

        var totalCharacters = 0;
        for (var index = 0; index < sections.Count; index++)
        {
            var section = sections[index];
            var path = $"{prefix}.sections[{index}]";
            ValidateText(section.Heading, $"{path}.heading", 1, 72);
            ValidateOptionalText(section.Body, $"{path}.body", 320);
            ValidateList(section.Bullets, $"{path}.bullets", 5, 100);
            ValidateOptionalText(section.Highlight, $"{path}.highlight", 48);
            if (string.IsNullOrWhiteSpace(section.Body)
                && (section.Bullets is null || section.Bullets.Count == 0))
            {
                throw new PptxValidationException(
                    "visual_content_missing",
                    $"{path} requires body or at least one bullet.");
            }

            if (!IsSupportedTone(section.Tone))
            {
                throw new PptxValidationException(
                    "visual_brief_section_tone_invalid",
                    $"{path}.tone must be a supported semantic tone or a #RRGGBB color.");
            }

            totalCharacters += section.Heading.Length
                + (section.Body?.Length ?? 0)
                + (section.Highlight?.Length ?? 0)
                + (section.Bullets?.Sum(static item => item.Length) ?? 0);
        }

        if (totalCharacters > 900)
        {
            throw new PptxValidationException(
                "visual_content_density_invalid",
                $"{prefix}.sections must not exceed 900 total characters; split the content across slides.");
        }
    }

    private static void ValidateScorecard(VisualScorecardSpec? scorecard, string prefix)
    {
        if (scorecard is null)
        {
            return;
        }

        if (scorecard.Options is null || scorecard.Options.Count is < 2 or > 4)
        {
            throw new PptxValidationException(
                "visual_scorecard_options_out_of_range",
                $"{prefix}.scorecard.options must contain between 2 and 4 options.");
        }

        if (scorecard.Criteria is null || scorecard.Criteria.Count is < 2 or > 6)
        {
            throw new PptxValidationException(
                "visual_scorecard_criteria_out_of_range",
                $"{prefix}.scorecard.criteria must contain between 2 and 6 rows.");
        }

        for (var optionIndex = 0; optionIndex < scorecard.Options.Count; optionIndex++)
        {
            var option = scorecard.Options[optionIndex];
            var path = $"{prefix}.scorecard.options[{optionIndex}]";
            ValidateText(option.Title, $"{path}.title", 1, 48);
            ValidateOptionalText(option.Subtitle, $"{path}.subtitle", 72);
        }

        for (var rowIndex = 0; rowIndex < scorecard.Criteria.Count; rowIndex++)
        {
            var row = scorecard.Criteria[rowIndex];
            var path = $"{prefix}.scorecard.criteria[{rowIndex}]";
            ValidateText(row.Criterion, $"{path}.criterion", 1, 48);
            if (row.Cells is null || row.Cells.Count != scorecard.Options.Count)
            {
                throw new PptxValidationException(
                    "visual_scorecard_cells_mismatch",
                    $"{path}.cells must contain exactly one cell per scorecard option.");
            }

            for (var cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                var cell = row.Cells[cellIndex];
                var cellPath = $"{path}.cells[{cellIndex}]";
                ValidateText(cell.Rating, $"{cellPath}.rating", 1, 24);
                ValidateOptionalText(cell.Detail, $"{cellPath}.detail", 100);
                if (!IsSupportedTone(cell.Tone))
                {
                    throw new PptxValidationException(
                        "visual_scorecard_tone_invalid",
                        $"{cellPath}.tone must be a supported semantic tone or a #RRGGBB color.");
                }
            }
        }
    }

    private static void ValidateMatrix(VisualMatrixSpec? matrix, string prefix)
    {
        if (matrix is null)
        {
            return;
        }

        ValidateText(matrix.HorizontalAxis, $"{prefix}.matrix.horizontalAxis", 1, 48);
        ValidateText(matrix.VerticalAxis, $"{prefix}.matrix.verticalAxis", 1, 48);
        if (matrix.Quadrants is null || matrix.Quadrants.Count != 4)
        {
            throw new PptxValidationException(
                "visual_matrix_invalid",
                $"{prefix}.matrix.quadrants must contain exactly four panels in top-left, top-right, bottom-left, bottom-right order.");
        }

        foreach (var quadrant in matrix.Quadrants)
        {
            ValidateText(quadrant.Title, $"{prefix}.matrix.quadrants.title", 1, 64);
            ValidateList(quadrant.Bullets, $"{prefix}.matrix.quadrants.bullets", 3, 80, minimumCount: 1);
            ValidateOptionalText(quadrant.Highlight, $"{prefix}.matrix.quadrants.highlight", 32);
        }
    }

    private static bool IsSupportedTone(string tone)
    {
        return !string.IsNullOrWhiteSpace(tone)
            && (SemanticTones.Contains(tone) || HexColorRegex().IsMatch(tone));
    }

    private static void ValidateChart(VisualChartSpec? chart, string prefix)
    {
        if (chart is null)
        {
            return;
        }

        ValidateList(chart.Categories, $"{prefix}.chart.categories", 12, 48, minimumCount: 2);
        ValidateOptionalText(chart.ValueSuffix, $"{prefix}.chart.valueSuffix", 12);
        if (chart.Series is null || chart.Series.Count is < 1 or > 4)
        {
            throw new PptxValidationException(
                "visual_chart_series_out_of_range",
                $"{prefix}.chart.series must contain between 1 and 4 series.");
        }

        foreach (var series in chart.Series)
        {
            ValidateText(series.Name, $"{prefix}.chart.series.name", 1, 64);
            if (series.Values is null || series.Values.Count != chart.Categories.Count)
            {
                throw new PptxValidationException(
                    "visual_chart_values_mismatch",
                    $"Every {prefix}.chart series must have one value per category.");
            }

            if (series.Values.Any(value => !double.IsFinite(value) || Math.Abs(value) > 1_000_000_000_000d))
            {
                throw new PptxValidationException(
                    "visual_chart_value_invalid",
                    $"{prefix}.chart values must be finite and within the supported range.");
            }
        }
    }

    private static void ValidateList(
        IReadOnlyList<string>? values,
        string path,
        int maximumCount,
        int maximumLength,
        int minimumCount = 0)
    {
        if (values is null)
        {
            if (minimumCount > 0)
            {
                throw new PptxValidationException("visual_content_missing", $"{path} is required.");
            }

            return;
        }

        if (values.Count < minimumCount || values.Count > maximumCount)
        {
            throw new PptxValidationException(
                "visual_content_density_invalid",
                $"{path} must contain between {minimumCount} and {maximumCount} items.");
        }

        for (var index = 0; index < values.Count; index++)
        {
            ValidateText(values[index], $"{path}[{index}]", 1, maximumLength);
        }
    }

    private static void RequireCount<T>(
        IReadOnlyList<T>? values,
        string prefix,
        string name,
        int minimum,
        int maximum)
    {
        if (values is null || values.Count < minimum || values.Count > maximum)
        {
            throw new PptxValidationException(
                "visual_content_missing",
                $"{prefix}.{name} must contain between {minimum} and {maximum} items for this layout.");
        }
    }

    private static void ValidateColor(string? value, string path)
    {
        if (value is not null && !HexColorRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "visual_color_invalid",
                $"{path} must be a six-digit RGB hex color, with an optional leading '#'.");
        }
    }

    private static void ValidateOptionalText(string? value, string path, int maximumLength)
    {
        if (value is not null)
        {
            ValidateText(value, path, 0, maximumLength);
        }
    }

    private static void ValidateText(string? value, string path, int minimumLength, int maximumLength)
    {
        if (value is null || value.Length < minimumLength || value.Length > maximumLength || ContainsInvalidControlCharacter(value))
        {
            throw new PptxValidationException(
                "visual_text_invalid",
                $"{path} must contain {minimumLength} to {maximumLength} characters and no control characters.");
        }
    }

    private static bool ContainsInvalidControlCharacter(string value) =>
        value.Any(character => char.IsControl(character) && character is not ('\n' or '\r' or '\t'));

    private static void ThrowDensity(string prefix, string name, int maximum) =>
        throw new PptxValidationException(
            "visual_content_density_invalid",
            $"{prefix}.{name} must not contain more than {maximum} items.");

    [GeneratedRegex("^#?[0-9A-Fa-f]{6}$", RegexOptions.CultureInvariant)]
    private static partial Regex HexColorRegex();

    [GeneratedRegex("^(?:[1-9][0-9]?)/(?:1|2|4|8|16)$", RegexOptions.CultureInvariant)]
    private static partial Regex MusicTimeSignatureRegex();

    [GeneratedRegex("^([A-Ga-g])([#b]?)([0-8])$", RegexOptions.CultureInvariant)]
    private static partial Regex MusicPitchRegex();
}
