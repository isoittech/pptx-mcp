using System.ComponentModel;
using System.Text.Json.Serialization;

namespace PptxMcp.Domain;

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualObjectArchetype>))]
public enum VisualObjectArchetype { Arrow, CurvedArrow, Frame, Callout, Bracket, Ring, Ribbon }

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualObjectPurpose>))]
public enum VisualObjectPurpose { Direction, Growth, Cycle, Emphasis, Grouping, Annotation }

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualObjectStyle>))]
public enum VisualObjectStyle { QuietCorporate, RoundedFriendly, Editorial, Technical }

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualObjectEmphasis>))]
public enum VisualObjectEmphasis { Subtle, Standard, Strong }

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualObjectOrientation>))]
public enum VisualObjectOrientation { Right, Left, Up, Down, Clockwise }

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualObjectPlacementRole>))]
public enum VisualObjectPlacementRole { HeaderAccent, ContentConnector, FocusFrame, ChartAnnotation, SectionDivider, BackgroundMotif }

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualObjectPaletteRole>))]
public enum VisualObjectPaletteRole { Primary, Secondary, Accent, Positive, Warning, Critical, Muted }

public sealed record VisualObjectBrief(
    [property: Description("Target slide number in the planned deck. It is audit metadata and does not start a draft.")]
    int SlideNumber,
    [property: Description("Semantic reason: direction, growth, cycle, emphasis, grouping, or annotation.")]
    VisualObjectPurpose VisualPurpose,
    [property: Description("Editable native PowerPoint archetype: arrow, curvedArrow, frame, callout, bracket, ring, or ribbon.")]
    VisualObjectArchetype Archetype,
    [property: Description("Reusable treatment: quietCorporate, roundedFriendly, editorial, or technical.")]
    VisualObjectStyle Style = VisualObjectStyle.QuietCorporate,
    [property: Description("Visual strength: subtle, standard, or strong. Prefer subtle and reserve strong for one focal object.")]
    VisualObjectEmphasis Emphasis = VisualObjectEmphasis.Subtle,
    [property: Description("Semantic direction: right, left, up, down, or clockwise.")]
    VisualObjectOrientation Orientation = VisualObjectOrientation.Right,
    [property: Description("Automatic placement: headerAccent, contentConnector, focusFrame, chartAnnotation, sectionDivider, or backgroundMotif.")]
    VisualObjectPlacementRole PlacementRole = VisualObjectPlacementRole.ContentConnector,
    [property: Description("Brand palette role only; raw colors are not accepted.")]
    VisualObjectPaletteRole PaletteRole = VisualObjectPaletteRole.Accent,
    [property: Description("Optional concise visible label, 1-48 characters.")]
    string? Label = null);

public sealed record VisualObjectAssetReference(
    [property: Description("Opaque asset ID returned by pptx_prepare_visual_objects in the same user and conversation scope.")]
    string AssetId);

public sealed record VisualObjectAssetManifest(
    string AssetId,
    string UserScope,
    string ConversationScope,
    VisualObjectBrief Brief,
    string Fingerprint,
    DateTimeOffset CreatedAt,
    DateTimeOffset ExpiresAt);

public sealed record VisualObjectRenderSpec(
    [property: JsonPropertyName("asset_id")] string AssetId,
    [property: JsonPropertyName("brief")] VisualObjectBrief Brief,
    [property: JsonPropertyName("fingerprint")] string Fingerprint);

public sealed record VisualObjectAssetView(
    [property: JsonPropertyName("asset_id")] string AssetId,
    [property: JsonPropertyName("slide_number")] int SlideNumber,
    [property: JsonPropertyName("archetype")] VisualObjectArchetype Archetype,
    [property: JsonPropertyName("placement_role")] VisualObjectPlacementRole PlacementRole,
    [property: JsonPropertyName("style")] VisualObjectStyle Style,
    [property: JsonPropertyName("emphasis")] VisualObjectEmphasis Emphasis,
    [property: JsonPropertyName("preview_description")] string PreviewDescription,
    [property: JsonPropertyName("expires_at")] DateTimeOffset ExpiresAt);

public sealed record VisualObjectPreparationView(
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("objects")] IReadOnlyList<VisualObjectAssetView> Objects,
    [property: JsonPropertyName("instruction")] string Instruction);

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<VisualDiagramKind>))]
public enum VisualDiagramKind { Tree, Flow, Cycle, Concentric, Network }

public sealed record VisualDiagramSpec(
    [property: Description("Editable semantic diagram: tree, flow, cycle, concentric, or network.")]
    VisualDiagramKind Kind,
    [property: Description("Two to twelve native-shape nodes. IDs are local opaque identifiers, not URLs or paths.")]
    IReadOnlyList<VisualDiagramNodeSpec> Nodes,
    [property: Description("Optional directed relationships. Tree/flow support up to 18 edges; cycle and concentric use node order.")]
    IReadOnlyList<VisualDiagramEdgeSpec>? Edges = null,
    [property: Description("Reading direction: leftToRight, topToBottom, or clockwise. The renderer chooses coordinates.")]
    string Direction = "leftToRight");

public sealed record VisualDiagramNodeSpec(
    [property: Description("Unique local node ID using ASCII letters, digits, hyphen, or underscore.")]
    string Id,
    [property: Description("Visible node title, 1-48 characters.")]
    string Label,
    [property: Description("Optional concise explanation, up to 100 characters.")]
    string? Description = null,
    [property: Description("Semantic tone or approved #RRGGBB color. Prefer neutral for most nodes and accent for one focal node.")]
    string Tone = "neutral",
    [property: Description("Whether this is a focal node. At most two nodes may be emphasized.")]
    bool Emphasize = false);

public sealed record VisualDiagramEdgeSpec(
    [property: Description("Source node ID.")] string From,
    [property: Description("Target node ID.")] string To,
    [property: Description("Optional concise relationship label, up to 32 characters.")] string? Label = null);
