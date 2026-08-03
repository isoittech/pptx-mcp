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

public sealed record TextReplacement(
    string Find,
    string Replace,
    int? SlideNumber = null,
    string? ShapeName = null,
    uint? ShapeId = null);

public sealed record TemplateField(
    int SlideNumber,
    string Text,
    string? ShapeName = null,
    uint? ShapeId = null);

public sealed record DeckField(
    string Text,
    string? ShapeName = null,
    uint? ShapeId = null,
    uint? PlaceholderIndex = null);

public sealed record DeckSlideSpec(
    string LayoutId,
    IReadOnlyList<DeckField> Fields);

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
    [property: JsonPropertyName("error_message")] string? ErrorMessage);
