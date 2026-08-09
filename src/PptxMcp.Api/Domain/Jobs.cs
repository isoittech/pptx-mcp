using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PptxMcp.Domain;

[JsonConverter(typeof(JsonStringEnumConverter<JobKind>))]
public enum JobKind
{
    Analyze,
    RenderPreview,
    ReplaceText,
    PopulateTemplate,
    CreateDeck,
    CreateVisualDeck,
    CreateBrandedVisualDeck,
}

[JsonConverter(typeof(JsonStringEnumConverter<JobState>))]
public enum JobState
{
    Queued,
    Running,
    Succeeded,
    Failed,
    Canceled,
}

[JsonConverter(typeof(JsonStringEnumConverter<DeckParagraphKind>))]
public enum DeckParagraphKind
{
    Plain,
    Bullet,
    Numbered,
}

public sealed record DeckParagraph(
    [property: JsonPropertyName("text"), Description("Text for one editable PowerPoint paragraph. Do not include a manual bullet or number prefix.")]
    string Text,
    [property: JsonPropertyName("kind"), Description("Paragraph semantics: Plain, Bullet, or Numbered.")]
    DeckParagraphKind Kind = DeckParagraphKind.Plain,
    [property: JsonPropertyName("level"), Description("Zero-based indentation level from 0 to 4.")]
    int Level = 0,
    [property: JsonPropertyName("start_at"), Description("Optional starting number from 1 to 32767. Only valid for Numbered paragraphs.")]
    int? StartAt = null);

public sealed record TextReplacement : IJsonOnDeserialized
{
    [JsonConstructor]
    public TextReplacement(
        string Find,
        string Replace,
        int? SlideNumber = null,
        string? ShapeName = null,
        uint? ShapeId = null)
    {
        this.Find = Find;
        this.Replace = Replace;
        this.SlideNumber = SlideNumber;
        this.ShapeName = ShapeName;
        this.ShapeId = ShapeId;
    }

    [JsonPropertyName("find")]
    public string Find { get; private set; }

    [JsonPropertyName("replace")]
    public string Replace { get; private set; }

    [JsonPropertyName("slide_number")]
    public int? SlideNumber { get; private set; }

    [JsonPropertyName("shape_name")]
    public string? ShapeName { get; private set; }

    [JsonPropertyName("shape_id")]
    public uint? ShapeId { get; private set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? LegacyProperties { get; set; }

    void IJsonOnDeserialized.OnDeserialized()
    {
        SlideNumber ??= LegacyJsonAliases.GetInt32(LegacyProperties, "slideNumber");
        ShapeName ??= LegacyJsonAliases.GetString(LegacyProperties, "shapeName");
        ShapeId ??= LegacyJsonAliases.GetUInt32(LegacyProperties, "shapeId");
        LegacyProperties = null;
    }
}

public sealed record TemplateField : IJsonOnDeserialized
{
    [JsonConstructor]
    public TemplateField(
        int SlideNumber,
        string? Text = null,
        string? ShapeName = null,
        uint? ShapeId = null,
        IReadOnlyList<DeckParagraph>? Paragraphs = null)
    {
        this.SlideNumber = SlideNumber;
        this.Text = Text;
        this.ShapeName = ShapeName;
        this.ShapeId = ShapeId;
        this.Paragraphs = Paragraphs;
    }

    [JsonPropertyName("slide_number")]
    public int SlideNumber { get; private set; }

    [JsonPropertyName("text")]
    public string? Text { get; private set; }

    [JsonPropertyName("paragraphs")]
    public IReadOnlyList<DeckParagraph>? Paragraphs { get; private set; }

    [JsonPropertyName("shape_name")]
    public string? ShapeName { get; private set; }

    [JsonPropertyName("shape_id")]
    public uint? ShapeId { get; private set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? LegacyProperties { get; set; }

    void IJsonOnDeserialized.OnDeserialized()
    {
        SlideNumber = LegacyJsonAliases.GetInt32(LegacyProperties, "slideNumber") ?? SlideNumber;
        ShapeName ??= LegacyJsonAliases.GetString(LegacyProperties, "shapeName");
        ShapeId ??= LegacyJsonAliases.GetUInt32(LegacyProperties, "shapeId");
        LegacyProperties = null;
    }
}

public sealed record DeckField : IJsonOnDeserialized
{
    [JsonConstructor]
    public DeckField(
        string? Text = null,
        string? ShapeName = null,
        uint? ShapeId = null,
        uint? PlaceholderIndex = null,
        IReadOnlyList<DeckParagraph>? Paragraphs = null)
    {
        this.Text = Text;
        this.ShapeName = ShapeName;
        this.ShapeId = ShapeId;
        this.PlaceholderIndex = PlaceholderIndex;
        this.Paragraphs = Paragraphs;
    }

    [JsonPropertyName("text"), Description("Plain text to place in the selected placeholder. Use paragraphs instead for real bullet or numbered lists.")]
    public string? Text { get; private set; }

    [JsonPropertyName("paragraphs"), Description("Editable PowerPoint paragraphs. Use one item per list entry and never type bullet glyphs or number prefixes into text.")]
    public IReadOnlyList<DeckParagraph>? Paragraphs { get; private set; }

    [JsonPropertyName("shape_name"), Description("Optional exact shape_name copied from pptx_analyze.")]
    public string? ShapeName { get; private set; }

    [JsonPropertyName("shape_id"), Description("Preferred exact shape_id copied from the selected layout in pptx_analyze.")]
    public uint? ShapeId { get; private set; }

    [JsonPropertyName("placeholder_index"), Description("Optional exact placeholder_index copied from pptx_analyze.")]
    public uint? PlaceholderIndex { get; private set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? LegacyProperties { get; set; }

    void IJsonOnDeserialized.OnDeserialized()
    {
        ShapeName ??= LegacyJsonAliases.GetString(LegacyProperties, "shapeName");
        ShapeId ??= LegacyJsonAliases.GetUInt32(LegacyProperties, "shapeId");
        PlaceholderIndex ??= LegacyJsonAliases.GetUInt32(LegacyProperties, "placeholderIndex");
        LegacyProperties = null;
    }
}

public sealed record DeckSlideSpec : IJsonOnDeserialized
{
    [JsonConstructor]
    public DeckSlideSpec(string LayoutId, IReadOnlyList<DeckField> Fields)
    {
        this.LayoutId = LayoutId ?? string.Empty;
        this.Fields = Fields ?? [];
    }

    [JsonPropertyName("layout_id"), Description("Exact layout_id copied verbatim from pptx_analyze. Never construct or modify this value.")]
    public string LayoutId { get; private set; }

    [JsonPropertyName("fields"), Description("Placeholder values for this slide. Use exactly one of text or paragraphs, plus shape_id whenever available.")]
    public IReadOnlyList<DeckField> Fields { get; private set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? LegacyProperties { get; set; }

    void IJsonOnDeserialized.OnDeserialized()
    {
        if (string.IsNullOrWhiteSpace(LayoutId))
        {
            LayoutId = LegacyJsonAliases.GetString(LegacyProperties, "layoutId") ?? string.Empty;
        }

        LegacyProperties = null;
    }
}

public sealed record DeckSlideRevision : IJsonOnDeserialized
{
    [JsonConstructor]
    public DeckSlideRevision(int SlideNumber, IReadOnlyList<DeckField> Fields)
    {
        this.SlideNumber = SlideNumber;
        this.Fields = Fields ?? [];
    }

    [JsonPropertyName("slide_number"), Description("One-based slide number from the successful deck job to revise.")]
    public int SlideNumber { get; private set; }

    [JsonPropertyName("fields"), Description("Complete replacement field list for this slide. Keep every field that must remain on the slide.")]
    public IReadOnlyList<DeckField> Fields { get; private set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? LegacyProperties { get; set; }

    void IJsonOnDeserialized.OnDeserialized()
    {
        SlideNumber = LegacyJsonAliases.GetInt32(LegacyProperties, "slideNumber") ?? SlideNumber;
        LegacyProperties = null;
    }
}

public sealed record ToolInputRequest(
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("tool")] string Tool,
    [property: JsonPropertyName("required_arguments")] IReadOnlyList<string> RequiredArguments,
    [property: JsonPropertyName("instruction")] string Instruction);

public sealed record ToolValidationError(
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("tool")] string Tool,
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("instruction")] string Instruction);

internal static class LegacyJsonAliases
{
    public static string? GetString(IReadOnlyDictionary<string, JsonElement>? properties, string name) =>
        properties is not null &&
        properties.TryGetValue(name, out var value) &&
        value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    public static int? GetInt32(IReadOnlyDictionary<string, JsonElement>? properties, string name) =>
        properties is not null &&
        properties.TryGetValue(name, out var value) &&
        value.TryGetInt32(out var result)
            ? result
            : null;

    public static uint? GetUInt32(IReadOnlyDictionary<string, JsonElement>? properties, string name) =>
        properties is not null &&
        properties.TryGetValue(name, out var value) &&
        value.TryGetUInt32(out var result)
            ? result
            : null;
}

public sealed record JobRecord
{
    public required string Id { get; init; }

    public required JobKind Kind { get; init; }

    public required JobState State { get; init; }

    public required string UserScope { get; init; }

    public required string ConversationScope { get; init; }

    public required string SourceFileId { get; init; }

    public required DateTimeOffset CreatedAt { get; init; }

    public required DateTimeOffset ExpiresAt { get; init; }

    public DateTimeOffset? StartedAt { get; init; }

    public DateTimeOffset? CompletedAt { get; init; }

    public DateTimeOffset? FirstDownloadedAt { get; init; }

    public int ProgressPercent { get; init; }

    public JsonElement? Payload { get; init; }

    public JsonElement? Result { get; init; }

    public IReadOnlyList<ArtifactRecord> Artifacts { get; init; } = [];

    public string? ErrorCode { get; init; }

    public string? ErrorMessage { get; init; }

    public string? ParentJobId { get; init; }

    public string? VisualRootJobId { get; init; }

    public int VisualRevisionRound { get; init; }

    public IReadOnlyList<int> VisualRevisedSlidesInRound { get; init; } = [];
}

public sealed record ArtifactRecord(
    string FileName,
    string MediaType,
    long Bytes,
    bool StartsDownloadRetention);

public sealed record PreviewImageData(
    int SlideNumber,
    string MediaType,
    byte[] Bytes);

public sealed record JobReceipt(
    [property: JsonPropertyName("job_id")] string JobId,
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("poll_after_seconds")] int PollAfterSeconds);

public sealed record ArtifactLink(
    [property: JsonPropertyName("file_name")] string FileName,
    [property: JsonPropertyName("media_type")] string MediaType,
    [property: JsonPropertyName("bytes")] long Bytes,
    [property: JsonPropertyName("url")] string Url,
    [property: JsonPropertyName("expires_at")] DateTimeOffset ExpiresAt);

public sealed record JobView(
    [property: JsonPropertyName("job_id")] string JobId,
    [property: JsonPropertyName("kind")] JobKind Kind,
    [property: JsonPropertyName("status")] JobState Status,
    [property: JsonPropertyName("progress_percent")] int ProgressPercent,
    [property: JsonPropertyName("created_at")] DateTimeOffset CreatedAt,
    [property: JsonPropertyName("completed_at")] DateTimeOffset? CompletedAt,
    [property: JsonPropertyName("result")] JsonElement? Result,
    [property: JsonPropertyName("artifacts")] IReadOnlyList<ArtifactLink> Artifacts,
    [property: JsonPropertyName("error_code")] string? ErrorCode,
    [property: JsonPropertyName("error_message")] string? ErrorMessage,
    [property: JsonPropertyName("visual_root_job_id")] string? VisualRootJobId,
    [property: JsonPropertyName("visual_revision_round")] int VisualRevisionRound,
    [property: JsonPropertyName("visual_revised_slides_in_round")] IReadOnlyList<int> VisualRevisedSlidesInRound);
