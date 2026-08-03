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
    string Language = "ja-JP");

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
    string? Takeaway = null);

public sealed record VisualMetricSpec(
    string Value,
    string Label,
    string? Detail = null,
    string Tone = "accent");

public sealed record VisualPanelSpec(
    string Title,
    IReadOnlyList<string> Bullets,
    string? Highlight = null);

public sealed record VisualStepSpec(
    string Title,
    string? Description = null,
    string? Label = null);

public sealed record VisualChartSpec(
    VisualChartKind Kind,
    IReadOnlyList<string> Categories,
    IReadOnlyList<VisualChartSeriesSpec> Series,
    string? ValueSuffix = null,
    bool ShowLegend = true);

public sealed record VisualChartSeriesSpec(
    string Name,
    IReadOnlyList<double> Values);

public sealed record VisualDeckCreationResult(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("layout_kinds")] IReadOnlyList<string> LayoutKinds,
    [property: JsonPropertyName("renderer")] string Renderer);

public static partial class VisualDeckValidator
{
    private static readonly HashSet<string> Presets = new(StringComparer.OrdinalIgnoreCase)
    {
        "midnight",
        "aurora",
        "sunset",
        "forest",
        "minimal",
    };

    private static readonly HashSet<string> MetricTones = new(StringComparer.OrdinalIgnoreCase)
    {
        "accent",
        "positive",
        "warning",
        "neutral",
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
        for (var index = 0; index < deck.Slides.Count; index++)
        {
            ValidateSlide(deck.Slides[index], index + 1);
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
                "theme.preset must be one of midnight, aurora, sunset, forest, or minimal.");
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
        ValidateList(slide.Bullets, $"{prefix}.bullets", 8, 180);

        if (slide.Metrics is not null)
        {
            if (slide.Metrics.Count > 4)
            {
                ThrowDensity(prefix, "metrics", 4);
            }

            foreach (var metric in slide.Metrics)
            {
                ValidateText(metric.Value, $"{prefix}.metrics.value", 1, 32);
                ValidateText(metric.Label, $"{prefix}.metrics.label", 1, 80);
                ValidateOptionalText(metric.Detail, $"{prefix}.metrics.detail", 140);
                if (!MetricTones.Contains(metric.Tone))
                {
                    throw new PptxValidationException(
                        "visual_metric_tone_invalid",
                        $"{prefix}.metrics.tone must be accent, positive, warning, or neutral.");
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
                RequireCount(slide.Metrics, prefix, "metrics", 2, 4);
                break;
            case VisualSlideKind.Comparison:
                RequireCount(slide.Panels, prefix, "panels", 2, 3);
                break;
            case VisualSlideKind.Process:
            case VisualSlideKind.Timeline:
                RequireCount(slide.Steps, prefix, "steps", 3, 6);
                break;
            case VisualSlideKind.Chart when slide.Chart is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.chart is required for a chart slide.");
            case VisualSlideKind.Quote when string.IsNullOrWhiteSpace(slide.Body):
                throw new PptxValidationException("visual_content_missing", $"{prefix}.body is required for a quote slide.");
        }
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
