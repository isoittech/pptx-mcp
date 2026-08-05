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
    string Variant = "auto");

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

    public static void Validate(VisualDeckSpec deck, int maximumSlides)
    {
        ArgumentNullException.ThrowIfNull(deck);
        ValidateText(deck.Title, "deck.title", 1, 160);
        ValidateOptionalText(deck.Subject, "deck.subject", 240);
        ValidateText(deck.Language, "deck.language", 2, 24);

        if (deck.Slides is null || deck.Slides.Count is < 1 || deck.Slides.Count > maximumSlides)
        {
            throw new PptxValidationException(
                "visual_slide_count_out_of_range",
                $"Visual decks must contain between 1 and {maximumSlides} slides.");
        }

        ValidateTheme(deck.Theme);
        ValidateDesign(deck.Design);
        for (var index = 0; index < deck.Slides.Count; index++)
        {
            ValidateSlide(deck.Slides[index], index + 1);
        }
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

        return warnings;
    }

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

        ValidateMatrix(slide.Matrix, prefix);

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
}
