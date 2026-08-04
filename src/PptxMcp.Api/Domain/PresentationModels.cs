using System.Text.Json.Serialization;

namespace PptxMcp.Domain;

public sealed record ShapeSummary(
    [property: JsonPropertyName("slide_number")] int SlideNumber,
    [property: JsonPropertyName("shape_id")] uint ShapeId,
    [property: JsonPropertyName("shape_name")] string ShapeName,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("text")] string? Text);

public sealed record SlideSummary(
    [property: JsonPropertyName("slide_number")] int SlideNumber,
    [property: JsonPropertyName("title")] string? Title,
    [property: JsonPropertyName("layout_name")] string? LayoutName,
    [property: JsonPropertyName("shapes")] IReadOnlyList<ShapeSummary> Shapes);

public sealed record PlaceholderSummary(
    [property: JsonPropertyName("shape_id")] uint ShapeId,
    [property: JsonPropertyName("shape_name")] string ShapeName,
    [property: JsonPropertyName("placeholder_type")] string PlaceholderType,
    [property: JsonPropertyName("placeholder_index")] uint? PlaceholderIndex);

public sealed record LayoutSummary(
    [property: JsonPropertyName("layout_id")] string LayoutId,
    [property: JsonPropertyName("layout_name")] string LayoutName,
    [property: JsonPropertyName("master_number")] int MasterNumber,
    [property: JsonPropertyName("placeholders")] IReadOnlyList<PlaceholderSummary> Placeholders);

public sealed record ThemeColorSummary(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("rgb")] string Rgb);

public sealed record PresentationThemeSummary(
    [property: JsonPropertyName("primary_color")] string? PrimaryColor,
    [property: JsonPropertyName("secondary_color")] string? SecondaryColor,
    [property: JsonPropertyName("accent_color")] string? AccentColor,
    [property: JsonPropertyName("background_color")] string? BackgroundColor,
    [property: JsonPropertyName("text_color")] string? TextColor,
    [property: JsonPropertyName("heading_font")] string? HeadingFont,
    [property: JsonPropertyName("body_font")] string? BodyFont,
    [property: JsonPropertyName("colors")] IReadOnlyList<ThemeColorSummary> Colors);

public sealed record PresentationSummary(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("has_smart_art")] bool HasSmartArt,
    [property: JsonPropertyName("has_charts")] bool HasCharts,
    [property: JsonPropertyName("has_embedded_workbook")] bool HasEmbeddedWorkbook,
    [property: JsonPropertyName("analysis_truncated")] bool AnalysisTruncated,
    [property: JsonPropertyName("slides")] IReadOnlyList<SlideSummary> Slides,
    [property: JsonPropertyName("layouts")] IReadOnlyList<LayoutSummary> Layouts,
    [property: JsonPropertyName("theme")] PresentationThemeSummary? Theme,
    [property: JsonPropertyName("validation_errors")] IReadOnlyList<string> ValidationErrors);

public sealed record EditResult(
    [property: JsonPropertyName("replacement_count")] int ReplacementCount,
    [property: JsonPropertyName("changed_parts")] IReadOnlyList<string> ChangedParts);

public sealed record DeckCreationResult(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("populated_field_count")] int PopulatedFieldCount,
    [property: JsonPropertyName("layout_ids")] IReadOnlyList<string> LayoutIds);
