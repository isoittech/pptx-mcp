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
    DataTable,
    Media,
    CoverageMap,
    TransformationEvidence,
    ArtifactShowcase,
    GanttSchedule,
    NativeDiagram,
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
    VisualDesignSpec? Design = null,
    [property: Description("Server-owned renderer contract. New staged decks use visual-v6-dom; saved visual-v4 and visual-v5 payloads retain their original renderer behavior during their stored lineage.")]
    string? RendererContract = null,
    [property: JsonPropertyName("brand_profile_binding"), Description("Server-owned immutable Brand Profile and per-slide recipe contract. Omitted for legacy and unprofiled decks.")]
    VisualDeckBrandProfileBinding? BrandProfileBinding = null,
    [property: JsonPropertyName("visual_object_assets"), Description("Server-owned immutable semantic visual object snapshots. Callers provide only slide.visualObjects asset IDs.")]
    IReadOnlyList<VisualObjectRenderSpec>? VisualObjectAssets = null);

public sealed record VisualDeckDraftView(
    [property: JsonPropertyName("draft_id")] string DraftId,
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("expected_slide_count")] int ExpectedSlideCount,
    [property: JsonPropertyName("accepted_slide_count")] int AcceptedSlideCount,
    [property: JsonPropertyName("next_slide_number")] int NextSlideNumber,
    [property: JsonPropertyName("remaining_slide_count")] int RemainingSlideCount,
    [property: JsonPropertyName("maximum_batch_slides")] int MaximumBatchSlides,
    [property: JsonPropertyName("template_source_file_id")] string TemplateSourceFileId,
    [property: JsonPropertyName("template_layout_id")] string TemplateLayoutId,
    [property: JsonPropertyName("creative_direction_locked")] bool CreativeDirectionLocked,
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
    [property: Description("Legacy single font for headings and body. Brand Profile workflows should prefer headingFontFace and bodyFontFace; omitted role-specific values fall back to this value.")]
    string? FontFace = null,
    [property: Description("Optional approved heading font. Omit to inherit fontFace or the renderer default.")]
    string? HeadingFontFace = null,
    [property: Description("Optional approved body font. Omit to inherit fontFace or the renderer default.")]
    string? BodyFontFace = null,
    [property: Description("Optional surface color for cards and tables as #RRGGBB.")]
    string? SurfaceColor = null,
    [property: Description("Optional muted text color as #RRGGBB.")]
    string? MutedTextColor = null,
    [property: Description("Optional positive semantic color as #RRGGBB.")]
    string? PositiveColor = null,
    [property: Description("Optional warning semantic color as #RRGGBB.")]
    string? WarningColor = null,
    [property: Description("Optional critical semantic color as #RRGGBB.")]
    string? CriticalColor = null,
    [property: Description("Optional ordered series palette of one to eight #RRGGBB colors for native charts.")]
    IReadOnlyList<string>? DataSeriesColors = null);

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
    [property: Description("Implemented layout variant. Use auto normally; split for Media or qualifying Bullets; spotlight for qualifying Metrics/Cards; editorial for a three-section StructuredBrief; loop for Process; stepped for Timeline/Roadmap; pyramid for Funnel. Other combinations are rejected rather than silently ignored.")]
    string Variant = "auto",
    [property: Description("Editable native standard-notation and ukulele TAB score for a music_score slide. Use semantic pitches, durations, strings, frets, and fingers instead of coordinates.")]
    VisualMusicScoreSpec? MusicScore = null,
    [property: Description("Editable general-purpose table for a data_table slide. Column headers and every row must use the same cell count.")]
    VisualDataTableSpec? DataTable = null,
    [property: Description("Verified user-uploaded image content for a media slide. The assetId is opaque and conversation-bound; placeholders, URLs, and paths are never valid substitutes.")]
    VisualMediaSpec? Media = null,
    [property: Description("Rows and bounded column spans for showing when policies, controls, or responsibilities apply. Use semantic groups and one-based column spans instead of coordinates.")]
    VisualCoverageMapSpec? CoverageMap = null,
    [property: Description("Editable before/after evidence with tagged input fragments, transformed output, and a verification table. Raw HTML is never accepted.")]
    VisualTransformationEvidenceSpec? TransformationEvidence = null,
    [property: Description("One to three groups of verified conversation-bound image assets presented as a matte document or deliverable showcase.")]
    VisualArtifactShowcaseSpec? ArtifactShowcase = null,
    [property: Description("Editable task schedule using bounded time columns and one-based start/end spans instead of arbitrary coordinates.")]
    VisualGanttScheduleSpec? GanttSchedule = null,
    [property: Description("Editable semantic diagram for a nativeDiagram slide. Coordinates, SVG, and XML are not accepted.")]
    VisualDiagramSpec? Diagram = null,
    [property: Description("Optional prepared native visual objects. Use at most three per slide and only opaque IDs returned by pptx_prepare_visual_objects in this conversation. A validated Design Brief materializes its planned IDs when omitted; any explicit list must match exactly.")]
    IReadOnlyList<VisualObjectAssetReference>? VisualObjects = null,
    [property: Description("Optional per-slide information density override: airy, balanced, or detailed. Omit to inherit deck design.density.")]
    string? Density = null,
    [property: Description("Optional immutable layout recipe ID selected from the active Brand Profile catalog. It never accepts coordinates, code, URLs, or paths.")]
    string? RecipeId = null,
    [property: Description("Optional PowerPoint speaker notes stored outside the visible slide canvas. For MSI-generated decks, provide both the slide purpose and presentation-ready talk script on every slide.")]
    VisualSpeakerNotesSpec? SpeakerNotes = null);

public sealed record VisualSpeakerNotesSpec(
    [property: Description("One concise sentence stating what this slide must communicate or persuade the audience to understand. This appears under the fixed 'このスライドの狙い' heading in PowerPoint speaker notes.")]
    string Purpose,
    [property: Description("Presentation-ready narration for this slide. Do not include hidden chain-of-thought, credentials, internal-only URLs, or content that recipients must not see.")]
    string TalkScript);

public sealed record VisualMediaSpec(
    [property: Description("Opaque asset_id returned by pptx_register_uploaded_image_asset in the same user and conversation scope.")]
    string AssetId,
    [property: Description("Crop intent: contain, cover, focalCenter, focalLeft, or focalRight.")]
    string CropIntent = "cover",
    [property: Description("Text column position for split media: left or right. The image occupies the opposite side.")]
    string TextPosition = "left",
    [property: Description("Optional short caption shown next to the image. The registered asset alt text remains the accessibility description.")]
    string? Caption = null);

public sealed record VisualChipSpec(
    [property: Description("Short visible label naming a category, state, phase, or approach.")]
    string Label,
    [property: Description("Semantic tone such as accent, positive, warning, critical, neutral, or a custom #RRGGBB color.")]
    string Tone = "neutral");

public sealed record VisualAxisColumnSpec(
    [property: Description("Stable lowercase identifier used only within this slide.")]
    string Id,
    [property: Description("Short visible column label such as W1, Planning, or Release.")]
    string Label,
    [property: Description("Optional contiguous parent heading such as a month or project phase.")]
    string? GroupLabel = null);

public sealed record VisualCoverageMapSpec(
    [property: Description("Two to six ordered columns that define the horizontal scope.")]
    IReadOnlyList<VisualAxisColumnSpec> Columns,
    [property: Description("One to four category groups containing the visible rows.")]
    IReadOnlyList<VisualCoverageGroupSpec> Groups,
    [property: Description("One to sixteen labeled spans. StartColumn and endColumn are one-based and inclusive.")]
    IReadOnlyList<VisualSpanBarSpec> Bars,
    [property: Description("Optional single practical takeaway pointing to a real row or bar ID.")]
    VisualCalloutSpec? Callout = null,
    [property: Description("Optional compact approach or status labels below the map.")]
    IReadOnlyList<VisualChipSpec>? FooterChips = null);

public sealed record VisualCoverageGroupSpec(
    string Id,
    string Label,
    string? Subtitle,
    [property: Description("One to five rows within this category.")]
    IReadOnlyList<VisualCoverageRowSpec> Rows,
    [property: Description("Semantic group color.")]
    string Tone = "accent");

public sealed record VisualCoverageRowSpec(
    string Id,
    string Label);

public sealed record VisualSpanBarSpec(
    [property: Description("Stable lowercase identifier used by an optional callout target.")]
    string Id,
    [property: Description("ID of a row declared in coverageMap.groups.rows.")]
    string RowId,
    string Label,
    [property: Description("One-based inclusive starting column.")]
    int StartColumn,
    [property: Description("One-based inclusive ending column.")]
    int EndColumn,
    string Tone = "accent");

public sealed record VisualCalloutSpec(
    string Text,
    [property: Description("Optional ID of the real row, bar, task, or evidence item being explained.")]
    string? TargetId = null,
    string Tone = "primary");

public sealed record VisualTransformationEvidenceSpec(
    string InputHeading,
    [property: Description("Ordered safe text fragments. Tags and tones are rendered by the server; HTML is not accepted.")]
    IReadOnlyList<VisualTaggedTextSegmentSpec> InputSegments,
    string OutputHeading,
    string OutputText,
    [property: Description("Editable table containing the evidence or detected entities supporting the transformation.")]
    VisualDataTableSpec EvidenceTable,
    string? InputCaption = null);

public sealed record VisualTaggedTextSegmentSpec(
    string Text,
    string? Tag = null,
    string Tone = "neutral");

public sealed record VisualArtifactShowcaseSpec(
    [property: Description("One to three deliverable or service groups.")]
    IReadOnlyList<VisualArtifactGroupSpec> Groups);

public sealed record VisualArtifactGroupSpec(
    string Title,
    [property: Description("One to four verified image assets shown with matte outlines and controlled overlap.")]
    IReadOnlyList<VisualArtifactItemSpec> Artifacts,
    string? Description = null);

public sealed record VisualArtifactItemSpec(
    [property: Description("Opaque asset_id returned by pptx_register_uploaded_image_asset in the same user and conversation scope.")]
    string AssetId,
    string? Label = null,
    [property: Description("Crop intent: contain or cover.")]
    string CropIntent = "contain");

public sealed record VisualGanttScheduleSpec(
    [property: Description("Four to twelve ordered time cells, optionally grouped by month or phase. Split longer schedules across slides so labels remain readable at 14pt or larger.")]
    IReadOnlyList<VisualAxisColumnSpec> Columns,
    [property: Description("Two to eight scheduled tasks in reading order. Split longer schedules across slides rather than shrinking text below 14pt.")]
    IReadOnlyList<VisualGanttTaskSpec> Tasks,
    string? EffortLabel = null,
    [property: Description("Optional labeled vertical ranges such as holidays, gates, or the current period.")]
    IReadOnlyList<VisualGanttMarkerSpec>? Markers = null,
    [property: Description("Optional compact legend defining every bar color used by the schedule.")]
    IReadOnlyList<VisualChipSpec>? Legend = null);

public sealed record VisualGanttTaskSpec(
    string Id,
    string Category,
    string Title,
    [property: Description("Up to three concise task details.")]
    IReadOnlyList<string>? Details,
    [property: Description("One-based inclusive starting column.")]
    int StartColumn,
    [property: Description("One-based inclusive ending column.")]
    int EndColumn,
    string Tone = "secondary");

public sealed record VisualGanttMarkerSpec(
    string Label,
    [property: Description("One-based inclusive starting column.")]
    int StartColumn,
    [property: Description("One-based inclusive ending column.")]
    int EndColumn,
    string Tone = "neutral");

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

public sealed record VisualDataTableSpec(
    [property: Description("Two to six editable table columns in reading order.")]
    IReadOnlyList<VisualDataTableColumnSpec> Columns,
    [property: Description("One to ten editable table rows. Every row must have exactly one cell per column.")]
    IReadOnlyList<VisualDataTableRowSpec> Rows,
    [property: Description("Style the first column as row headings. Enabled by default for scannable business tables.")]
    bool FirstColumnIsHeader = true);

public sealed record VisualDataTableColumnSpec(
    [property: Description("Short visible column heading.")]
    string Header,
    [property: Description("Cell alignment for this column: left, center, or right.")]
    string Align = "left",
    [property: Description("Relative width from 0.5 through 4.0. Values are normalized across the available table width.")]
    double WidthWeight = 1);

public sealed record VisualDataTableRowSpec(
    [property: Description("One cell per table column, in exactly the same order as dataTable.columns.")]
    IReadOnlyList<VisualDataTableCellSpec> Cells);

public sealed record VisualDataTableCellSpec(
    [property: Description("Concise visible cell text.")]
    string Text,
    [property: Description("Semantic tone such as positive, warning, critical, neutral, or a custom #RRGGBB color.")]
    string Tone = "neutral",
    [property: Description("Use restrained bold emphasis for this cell.")]
    bool Emphasize = false);

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
    [property: Description("Editable server-approved Lucide icon ID. Common choices include insight, target, growth, people, person, organization, shield, clock, calendar, gantt, workflow, layers, files, table, flag, brain, sparkles, gauge, route, presentation, training, business, checklist, scan, tags, components, cloud, settings, data, warning, check, search, compliance, decision, lock, network, document, communication, recovery, backup, legal, monitor, or automation.")]
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
            PrimaryColor = current.PrimaryColor ?? templateTheme.PrimaryColor,
            SecondaryColor = current.SecondaryColor ?? templateTheme.SecondaryColor,
            AccentColor = current.AccentColor ?? templateTheme.AccentColor,
            BackgroundColor = current.BackgroundColor ?? templateTheme.BackgroundColor,
            TextColor = current.TextColor ?? templateTheme.TextColor,
            FontFace = current.FontFace ?? templateTheme.BodyFont ?? templateTheme.HeadingFont,
            HeadingFontFace = current.HeadingFontFace
                ?? current.FontFace
                ?? templateTheme.HeadingFont
                ?? templateTheme.BodyFont,
            BodyFontFace = current.BodyFontFace
                ?? current.FontFace
                ?? templateTheme.BodyFont
                ?? templateTheme.HeadingFont,
        };
        return deck with { Theme = brandedTheme };
    }
}

public sealed record VisualDeckCreationResult(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("layout_kinds")] IReadOnlyList<string> LayoutKinds,
    [property: JsonPropertyName("renderer")] string Renderer,
    [property: JsonPropertyName("speaker_notes_count")] int SpeakerNotesCount,
    [property: JsonPropertyName("design_warnings")] IReadOnlyList<string> DesignWarnings);

public sealed record BrandedVisualDeckCreationResult(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("layout_kinds")] IReadOnlyList<string> LayoutKinds,
    [property: JsonPropertyName("renderer")] string Renderer,
    [property: JsonPropertyName("template_layout_id")] string TemplateLayoutId,
    [property: JsonPropertyName("template_layout_name")] string TemplateLayoutName,
    [property: JsonPropertyName("template_theme_applied")] bool TemplateThemeApplied,
    [property: JsonPropertyName("speaker_notes_count")] int SpeakerNotesCount,
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
        "brain",
        "workflow",
        "layers",
        "files",
        "table",
        "calendar",
        "gantt",
        "flag",
        "person",
        "organization",
        "sparkles",
        "gauge",
        "route",
        "arrow",
        "presentation",
        "training",
        "business",
        "checklist",
        "scan",
        "tags",
        "components",
    };

    private static readonly HashSet<string> MediaCropIntents = new(StringComparer.OrdinalIgnoreCase)
    {
        "contain",
        "cover",
        "focalCenter",
        "focalLeft",
        "focalRight",
    };

    private static readonly HashSet<string> MediaTextPositions = new(StringComparer.OrdinalIgnoreCase)
    {
        "left",
        "right",
    };

    private static readonly HashSet<string> SlideVariants = new(StringComparer.OrdinalIgnoreCase)
    {
        "auto",
        "spotlight",
        "split",
        "editorial",
        "stepped",
        "pyramid",
        "loop",
    };

    private static readonly HashSet<string> LegacySlideVariants = new(StringComparer.OrdinalIgnoreCase)
    {
        "auto",
        "grid",
        "spotlight",
        "split",
        "cascade",
        "editorial",
    };

    private static readonly HashSet<string> TableAlignments = new(StringComparer.OrdinalIgnoreCase)
    {
        "left",
        "center",
        "right",
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
        var rendererContract = deck.RendererContract?.ToLowerInvariant() ?? "visual-v4";
        if (rendererContract is not ("visual-v4" or "visual-v5" or "visual-v6-dom"))
        {
            throw new PptxValidationException(
                "visual_renderer_contract_invalid",
                "rendererContract is server-owned and must be visual-v4, visual-v5, or visual-v6-dom.");
        }

        if (deck.Slides is null || deck.Slides.Count is < 1 || deck.Slides.Count > maximumSlides)
        {
            throw new PptxValidationException(
                "visual_slide_count_out_of_range",
                $"Visual decks must contain between 1 and {maximumSlides} slides.");
        }

        for (var index = 0; index < deck.Slides.Count; index++)
        {
            ValidateSlide(
                deck.Slides[index],
                index + 1,
                deck.Design?.Density,
                rendererContract is "visual-v5" or "visual-v6-dom");
        }

        ValidateVisualObjectAssets(deck);
    }

    private static void ValidateVisualObjectAssets(VisualDeckSpec deck)
    {
        var references = deck.Slides
            .SelectMany(static slide => slide.VisualObjects ?? [])
            .Select(static item => item.AssetId)
            .ToArray();
        if (references.Length == 0)
        {
            if (deck.VisualObjectAssets is { Count: > 0 })
            {
                throw new PptxValidationException(
                    "visual_object_assets_unreferenced",
                    "Server-owned visual_object_assets must not contain unreferenced objects.");
            }

            return;
        }

        if (references.Length > VisualObjectAssetRepository.MaximumConversationObjects
            || deck.VisualObjectAssets is null
            || deck.VisualObjectAssets.Count != references.Distinct(StringComparer.Ordinal).Count())
        {
            throw new PptxValidationException(
                "visual_object_assets_missing",
                "Every slide.visualObjects asset ID must have exactly one server-owned semantic snapshot.");
        }

        var assets = new HashSet<string>(StringComparer.Ordinal);
        foreach (var asset in deck.VisualObjectAssets)
        {
            if (asset is null
                || !ImageAssetIdRegex().IsMatch(asset.AssetId)
                || !assets.Add(asset.AssetId)
                || asset.Fingerprint.Length != 64
                || !asset.Fingerprint.All(static character => character is >= '0' and <= '9' or >= 'a' and <= 'f'))
            {
                throw new PptxValidationException(
                    "visual_object_assets_invalid",
                    "Server-owned visual_object_assets contain an invalid ID or fingerprint.");
            }
        }

        if (references.Any(reference => !assets.Contains(reference)))
        {
            throw new PptxValidationException(
                "visual_object_assets_missing",
                "Every slide.visualObjects asset ID must have exactly one server-owned semantic snapshot.");
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

        for (var index = 0; index < deck.Slides.Count; index++)
        {
            var slide = deck.Slides[index];
            var usesDetailedDensity = string.Equals(
                slide.Density ?? deck.Design?.Density,
                "detailed",
                StringComparison.OrdinalIgnoreCase);
            if (GetTextCharacterCount(slide) >= 500 && !usesDetailedDensity)
            {
                warnings.Add($"Slide {index + 1} is text-rich; use slide density=detailed or deck design.density=detailed and a structured_brief, scorecard, or data_table layout instead of shrinking text.");
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
            + (slide.Media?.Caption?.Length ?? 0)
            + (slide.CoverageMap?.Columns.Sum(static item => item.Label.Length + (item.GroupLabel?.Length ?? 0)) ?? 0)
            + (slide.CoverageMap?.Groups.Sum(static item => item.Label.Length + (item.Subtitle?.Length ?? 0) + item.Rows.Sum(static row => row.Label.Length)) ?? 0)
            + (slide.CoverageMap?.Bars.Sum(static item => item.Label.Length) ?? 0)
            + (slide.CoverageMap?.Callout?.Text.Length ?? 0)
            + (slide.CoverageMap?.FooterChips?.Sum(static item => item.Label.Length) ?? 0)
            + (slide.TransformationEvidence?.InputHeading.Length ?? 0)
            + (slide.TransformationEvidence?.InputCaption?.Length ?? 0)
            + (slide.TransformationEvidence?.InputSegments.Sum(static item => item.Text.Length + (item.Tag?.Length ?? 0)) ?? 0)
            + (slide.TransformationEvidence?.OutputHeading.Length ?? 0)
            + (slide.TransformationEvidence?.OutputText.Length ?? 0)
            + (slide.ArtifactShowcase?.Groups.Sum(static item => item.Title.Length + (item.Description?.Length ?? 0) + item.Artifacts.Sum(static artifact => artifact.Label?.Length ?? 0)) ?? 0)
            + (slide.GanttSchedule?.Columns.Sum(static item => item.Label.Length + (item.GroupLabel?.Length ?? 0)) ?? 0)
            + (slide.GanttSchedule?.Tasks.Sum(static item => item.Category.Length + item.Title.Length + (item.Details?.Sum(static detail => detail.Length) ?? 0)) ?? 0)
            + (slide.GanttSchedule?.EffortLabel?.Length ?? 0)
            + (slide.GanttSchedule?.Markers?.Sum(static item => item.Label.Length) ?? 0)
            + (slide.GanttSchedule?.Legend?.Sum(static item => item.Label.Length) ?? 0)
            + (slide.Diagram?.Nodes.Sum(static node => node.Label.Length + (node.Description?.Length ?? 0)) ?? 0)
            + (slide.Diagram?.Edges?.Sum(static edge => edge.Label?.Length ?? 0) ?? 0)
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

        if (slide.DataTable is not null)
        {
            count += slide.DataTable.Columns.Sum(static column => column.Header.Length);
            count += slide.DataTable.Rows.Sum(static row =>
                row.Cells.Sum(static cell => cell.Text.Length));
        }

        if (slide.TransformationEvidence?.EvidenceTable is { } evidenceTable)
        {
            count += evidenceTable.Columns.Sum(static column => column.Header.Length);
            count += evidenceTable.Rows.Sum(static row =>
                row.Cells.Sum(static cell => cell.Text.Length));
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
        ValidateOptionalText(theme.HeadingFontFace, "theme.headingFontFace", 80);
        ValidateOptionalText(theme.BodyFontFace, "theme.bodyFontFace", 80);
        ValidateColor(theme.SurfaceColor, "theme.surfaceColor");
        ValidateColor(theme.MutedTextColor, "theme.mutedTextColor");
        ValidateColor(theme.PositiveColor, "theme.positiveColor");
        ValidateColor(theme.WarningColor, "theme.warningColor");
        ValidateColor(theme.CriticalColor, "theme.criticalColor");
        ValidateOptionalContrast(
            theme.TextColor,
            theme.SurfaceColor,
            4.5,
            "theme.textColor",
            "theme.surfaceColor");
        ValidateOptionalContrast(
            theme.MutedTextColor,
            theme.SurfaceColor,
            3,
            "theme.mutedTextColor",
            "theme.surfaceColor");
        if (theme.DataSeriesColors is not null)
        {
            if (theme.DataSeriesColors.Count is < 1 or > 8)
            {
                throw new PptxValidationException(
                    "visual_theme_series_invalid",
                    "theme.dataSeriesColors must contain between one and eight colors when specified.");
            }

            for (var index = 0; index < theme.DataSeriesColors.Count; index++)
            {
                ValidateColor(theme.DataSeriesColors[index], $"theme.dataSeriesColors[{index}]");
            }
        }
    }

    private static void ValidateSlide(
        VisualSlideSpec slide,
        int slideNumber,
        string? deckDensity,
        bool usesModernRendererContract)
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
        ValidateOptionalText(slide.RecipeId, $"{prefix}.recipeId", 128);
        ValidateSpeakerNotes(slide.SpeakerNotes, prefix);
        if (slide.RecipeId is not null && !OpaqueIdentifierRegex().IsMatch(slide.RecipeId))
        {
            throw new PptxValidationException(
                "visual_slide_recipe_invalid",
                $"{prefix}.recipeId may contain only ASCII letters, digits, hyphens, and underscores.");
        }

        if (slide.Density is not null && !DesignDensities.Contains(slide.Density))
        {
            throw new PptxValidationException(
                "visual_slide_density_invalid",
                $"{prefix}.density must be airy, balanced, or detailed when specified.");
        }

        var acceptedVariants = usesModernRendererContract ? SlideVariants : LegacySlideVariants;
        if (!acceptedVariants.Contains(slide.Variant))
        {
            throw new PptxValidationException(
                "visual_slide_variant_invalid",
                usesModernRendererContract
                    ? $"{prefix}.variant must be auto, spotlight, split, editorial, stepped, pyramid, or loop."
                    : $"{prefix}.variant is not recognized by the stored legacy renderer contract.");
        }

        if (usesModernRendererContract)
        {
            ValidateVariantForSlide(slide, prefix);
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
        ValidateDataTable(
            slide.DataTable,
            prefix,
            slide.Density ?? deckDensity ?? "balanced",
            !string.IsNullOrWhiteSpace(slide.Takeaway));
        ValidateMatrix(slide.Matrix, prefix);
        ValidateMedia(slide.Media, prefix);
        ValidateCoverageMap(slide.CoverageMap, prefix);
        ValidateTransformationEvidence(
            slide.TransformationEvidence,
            prefix,
            slide.Density ?? deckDensity ?? "balanced");
        ValidateArtifactShowcase(slide.ArtifactShowcase, prefix);
        ValidateGanttSchedule(slide.GanttSchedule, prefix);
        ValidateDiagram(slide.Diagram, prefix);
        ValidateVisualObjectReferences(slide.VisualObjects, prefix);
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
            case VisualSlideKind.DataTable when slide.DataTable is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.dataTable is required for a dataTable slide.");
            case VisualSlideKind.Media when slide.Media is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.media with a verified assetId is required for a Media slide; placeholders are not accepted.");
            case VisualSlideKind.CoverageMap when slide.CoverageMap is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.coverageMap is required for a CoverageMap slide.");
            case VisualSlideKind.TransformationEvidence when slide.TransformationEvidence is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.transformationEvidence is required for a TransformationEvidence slide.");
            case VisualSlideKind.ArtifactShowcase when slide.ArtifactShowcase is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.artifactShowcase with verified asset IDs is required for an ArtifactShowcase slide.");
            case VisualSlideKind.GanttSchedule when slide.GanttSchedule is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.ganttSchedule is required for a GanttSchedule slide.");
            case VisualSlideKind.NativeDiagram when slide.Diagram is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.diagram is required for a NativeDiagram slide.");
            case VisualSlideKind.MusicScore when slide.MusicScore is null:
                throw new PptxValidationException("visual_content_missing", $"{prefix}.musicScore is required for a musicScore slide.");
        }
    }

    private static void ValidateCoverageMap(VisualCoverageMapSpec? coverageMap, string prefix)
    {
        if (coverageMap is null)
        {
            return;
        }

        var path = $"{prefix}.coverageMap";
        ValidateAxisColumns(coverageMap.Columns, path, 2, 6);
        if (coverageMap.Groups is null || coverageMap.Groups.Count is < 1 or > 4)
        {
            throw new PptxValidationException(
                "visual_coverage_groups_out_of_range",
                $"{path}.groups must contain between 1 and 4 groups.");
        }

        var localIds = new HashSet<string>(StringComparer.Ordinal);
        var rowIds = new HashSet<string>(StringComparer.Ordinal);
        var rowCount = 0;
        for (var groupIndex = 0; groupIndex < coverageMap.Groups.Count; groupIndex++)
        {
            var group = coverageMap.Groups[groupIndex]
                ?? throw new PptxValidationException("visual_coverage_group_invalid", $"{path}.groups[{groupIndex}] must not be null.");
            var groupPath = $"{path}.groups[{groupIndex}]";
            ValidateLocalId(group.Id, $"{groupPath}.id", localIds);
            ValidateText(group.Label, $"{groupPath}.label", 1, 40);
            ValidateOptionalText(group.Subtitle, $"{groupPath}.subtitle", 40);
            ValidateTone(group.Tone, $"{groupPath}.tone");
            if (group.Rows is null || group.Rows.Count is < 1 or > 5)
            {
                throw new PptxValidationException(
                    "visual_coverage_rows_out_of_range",
                    $"{groupPath}.rows must contain between 1 and 5 rows.");
            }

            rowCount += group.Rows.Count;
            for (var rowIndex = 0; rowIndex < group.Rows.Count; rowIndex++)
            {
                var row = group.Rows[rowIndex]
                    ?? throw new PptxValidationException("visual_coverage_row_invalid", $"{groupPath}.rows[{rowIndex}] must not be null.");
                var rowPath = $"{groupPath}.rows[{rowIndex}]";
                ValidateLocalId(row.Id, $"{rowPath}.id", localIds);
                rowIds.Add(row.Id);
                ValidateText(row.Label, $"{rowPath}.label", 1, 48);
            }
        }

        if (rowCount > 8)
        {
            throw new PptxValidationException(
                "visual_coverage_rows_out_of_range",
                $"{path} may contain at most 8 rows; split the map across slides so text remains at least 14pt.");
        }

        if (coverageMap.Bars is null || coverageMap.Bars.Count is < 1 or > 16)
        {
            throw new PptxValidationException(
                "visual_coverage_bars_out_of_range",
                $"{path}.bars must contain between 1 and 16 spans.");
        }

        var barIds = new HashSet<string>(StringComparer.Ordinal);
        for (var barIndex = 0; barIndex < coverageMap.Bars.Count; barIndex++)
        {
            var bar = coverageMap.Bars[barIndex]
                ?? throw new PptxValidationException("visual_coverage_bar_invalid", $"{path}.bars[{barIndex}] must not be null.");
            var barPath = $"{path}.bars[{barIndex}]";
            ValidateLocalId(bar.Id, $"{barPath}.id", localIds);
            barIds.Add(bar.Id);
            if (!rowIds.Contains(bar.RowId))
            {
                throw new PptxValidationException(
                    "visual_coverage_row_reference_invalid",
                    $"{barPath}.rowId must reference a row declared in {path}.groups.");
            }

            ValidateText(bar.Label, $"{barPath}.label", 1, 80);
            ValidateSpan(bar.StartColumn, bar.EndColumn, coverageMap.Columns.Count, barPath);
            ValidateTone(bar.Tone, $"{barPath}.tone");
        }

        if (coverageMap.Callout is { } callout)
        {
            ValidateText(callout.Text, $"{path}.callout.text", 1, 180);
            ValidateTone(callout.Tone, $"{path}.callout.tone");
            if (callout.TargetId is not null
                && !rowIds.Contains(callout.TargetId)
                && !barIds.Contains(callout.TargetId))
            {
                throw new PptxValidationException(
                    "visual_callout_target_invalid",
                    $"{path}.callout.targetId must reference a real row or bar ID on the same slide.");
            }
        }

        ValidateChips(coverageMap.FooterChips, $"{path}.footerChips", 6);
    }

    private static void ValidateTransformationEvidence(
        VisualTransformationEvidenceSpec? evidence,
        string prefix,
        string density)
    {
        if (evidence is null)
        {
            return;
        }

        var path = $"{prefix}.transformationEvidence";
        ValidateText(evidence.InputHeading, $"{path}.inputHeading", 1, 56);
        ValidateOptionalText(evidence.InputCaption, $"{path}.inputCaption", 80);
        ValidateText(evidence.OutputHeading, $"{path}.outputHeading", 1, 56);
        ValidateText(evidence.OutputText, $"{path}.outputText", 1, 1_200);
        if (evidence.InputSegments is null || evidence.InputSegments.Count is < 1 or > 40)
        {
            throw new PptxValidationException(
                "visual_transformation_segments_out_of_range",
                $"{path}.inputSegments must contain between 1 and 40 ordered fragments.");
        }

        var totalCharacters = evidence.OutputText.Length;
        for (var index = 0; index < evidence.InputSegments.Count; index++)
        {
            var segment = evidence.InputSegments[index]
                ?? throw new PptxValidationException("visual_transformation_segment_invalid", $"{path}.inputSegments[{index}] must not be null.");
            var segmentPath = $"{path}.inputSegments[{index}]";
            ValidateText(segment.Text, $"{segmentPath}.text", 1, 240);
            ValidateOptionalText(segment.Tag, $"{segmentPath}.tag", 24);
            ValidateTone(segment.Tone, $"{segmentPath}.tone");
            totalCharacters += segment.Text.Length + (segment.Tag?.Length ?? 0);
        }

        if (totalCharacters > 1_800)
        {
            throw new PptxValidationException(
                "visual_content_density_invalid",
                $"{path} must not exceed 1800 input/output characters; split the evidence across slides.");
        }

        ValidateDataTable(evidence.EvidenceTable, $"{path}.evidenceTable", density, false);
    }

    private static void ValidateArtifactShowcase(VisualArtifactShowcaseSpec? showcase, string prefix)
    {
        if (showcase is null)
        {
            return;
        }

        var path = $"{prefix}.artifactShowcase";
        if (showcase.Groups is null || showcase.Groups.Count is < 1 or > 3)
        {
            throw new PptxValidationException(
                "visual_artifact_groups_out_of_range",
                $"{path}.groups must contain between 1 and 3 groups.");
        }

        var assetIds = new HashSet<string>(StringComparer.Ordinal);
        for (var groupIndex = 0; groupIndex < showcase.Groups.Count; groupIndex++)
        {
            var group = showcase.Groups[groupIndex]
                ?? throw new PptxValidationException("visual_artifact_group_invalid", $"{path}.groups[{groupIndex}] must not be null.");
            var groupPath = $"{path}.groups[{groupIndex}]";
            ValidateText(group.Title, $"{groupPath}.title", 1, 72);
            ValidateOptionalText(group.Description, $"{groupPath}.description", 120);
            if (group.Artifacts is null || group.Artifacts.Count is < 1 or > 4)
            {
                throw new PptxValidationException(
                    "visual_artifacts_out_of_range",
                    $"{groupPath}.artifacts must contain between 1 and 4 verified assets.");
            }

            for (var artifactIndex = 0; artifactIndex < group.Artifacts.Count; artifactIndex++)
            {
                var artifact = group.Artifacts[artifactIndex]
                    ?? throw new PptxValidationException("visual_artifact_invalid", $"{groupPath}.artifacts[{artifactIndex}] must not be null.");
                var artifactPath = $"{groupPath}.artifacts[{artifactIndex}]";
                if (!ImageAssetIdRegex().IsMatch(artifact.AssetId) || !assetIds.Add(artifact.AssetId))
                {
                    throw new PptxValidationException(
                        "visual_artifact_asset_id_invalid",
                        $"{artifactPath}.assetId must be a unique opaque asset_id returned by pptx_register_uploaded_image_asset.");
                }

                ValidateOptionalText(artifact.Label, $"{artifactPath}.label", 56);
                if (artifact.CropIntent is not ("contain" or "cover"))
                {
                    throw new PptxValidationException(
                        "visual_artifact_crop_invalid",
                        $"{artifactPath}.cropIntent must be contain or cover.");
                }
            }
        }
    }

    private static void ValidateGanttSchedule(VisualGanttScheduleSpec? gantt, string prefix)
    {
        if (gantt is null)
        {
            return;
        }

        var path = $"{prefix}.ganttSchedule";
        ValidateAxisColumns(gantt.Columns, path, 4, 12);
        ValidateOptionalText(gantt.EffortLabel, $"{path}.effortLabel", 56);
        if (gantt.Tasks is null || gantt.Tasks.Count is < 2 or > 8)
        {
            throw new PptxValidationException(
                "visual_gantt_tasks_out_of_range",
                $"{path}.tasks must contain between 2 and 8 tasks; split longer schedules so text remains at least 14pt.");
        }

        var taskIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < gantt.Tasks.Count; index++)
        {
            var task = gantt.Tasks[index]
                ?? throw new PptxValidationException("visual_gantt_task_invalid", $"{path}.tasks[{index}] must not be null.");
            var taskPath = $"{path}.tasks[{index}]";
            ValidateLocalId(task.Id, $"{taskPath}.id", taskIds);
            ValidateText(task.Category, $"{taskPath}.category", 1, 40);
            ValidateText(task.Title, $"{taskPath}.title", 1, 90);
            ValidateList(task.Details, $"{taskPath}.details", 3, 100);
            ValidateSpan(task.StartColumn, task.EndColumn, gantt.Columns.Count, taskPath);
            ValidateTone(task.Tone, $"{taskPath}.tone");
        }

        if (gantt.Markers is { Count: > 3 })
        {
            throw new PptxValidationException(
                "visual_gantt_markers_out_of_range",
                $"{path}.markers may contain at most 3 labeled ranges.");
        }

        if (gantt.Markers is not null)
        {
            for (var index = 0; index < gantt.Markers.Count; index++)
            {
                var marker = gantt.Markers[index]
                    ?? throw new PptxValidationException("visual_gantt_marker_invalid", $"{path}.markers[{index}] must not be null.");
                var markerPath = $"{path}.markers[{index}]";
                ValidateText(marker.Label, $"{markerPath}.label", 1, 40);
                ValidateSpan(marker.StartColumn, marker.EndColumn, gantt.Columns.Count, markerPath);
                ValidateTone(marker.Tone, $"{markerPath}.tone");
            }
        }

        ValidateChips(gantt.Legend, $"{path}.legend", 6);
    }

    private static void ValidateAxisColumns(
        IReadOnlyList<VisualAxisColumnSpec>? columns,
        string path,
        int minimumCount,
        int maximumCount)
    {
        if (columns is null || columns.Count < minimumCount || columns.Count > maximumCount)
        {
            throw new PptxValidationException(
                "visual_axis_columns_out_of_range",
                $"{path}.columns must contain between {minimumCount} and {maximumCount} columns.");
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < columns.Count; index++)
        {
            var column = columns[index]
                ?? throw new PptxValidationException("visual_axis_column_invalid", $"{path}.columns[{index}] must not be null.");
            var columnPath = $"{path}.columns[{index}]";
            ValidateLocalId(column.Id, $"{columnPath}.id", ids);
            ValidateSingleLineText(column.Label, $"{columnPath}.label", 1, 32);
            ValidateOptionalText(column.GroupLabel, $"{columnPath}.groupLabel", 40);
        }
    }

    private static void ValidateChips(IReadOnlyList<VisualChipSpec>? chips, string path, int maximumCount)
    {
        if (chips is null)
        {
            return;
        }

        if (chips.Count > maximumCount)
        {
            throw new PptxValidationException(
                "visual_chips_out_of_range",
                $"{path} may contain at most {maximumCount} chips.");
        }

        for (var index = 0; index < chips.Count; index++)
        {
            var chip = chips[index]
                ?? throw new PptxValidationException("visual_chip_invalid", $"{path}[{index}] must not be null.");
            ValidateSingleLineText(chip.Label, $"{path}[{index}].label", 1, 40);
            ValidateTone(chip.Tone, $"{path}[{index}].tone");
        }
    }

    private static void ValidateSpan(int startColumn, int endColumn, int columnCount, string path)
    {
        if (startColumn < 1 || endColumn < startColumn || endColumn > columnCount)
        {
            throw new PptxValidationException(
                "visual_span_invalid",
                $"{path} must use a one-based inclusive startColumn/endColumn within the declared columns.");
        }
    }

    private static void ValidateTone(string tone, string path)
    {
        if (!IsSupportedTone(tone))
        {
            throw new PptxValidationException(
                "visual_tone_invalid",
                $"{path} must be a supported semantic tone or a #RRGGBB color.");
        }
    }

    private static void ValidateLocalId(string value, string path, HashSet<string> existingIds)
    {
        if (!LocalVisualIdRegex().IsMatch(value) || !existingIds.Add(value))
        {
            throw new PptxValidationException(
                "visual_local_id_invalid",
                $"{path} must be unique on the slide and contain only lowercase letters, digits, hyphens, or underscores.");
        }
    }

    private static void ValidateMedia(VisualMediaSpec? media, string prefix)
    {
        if (media is null)
        {
            return;
        }

        var path = $"{prefix}.media";
        if (!ImageAssetIdRegex().IsMatch(media.AssetId))
        {
            throw new PptxValidationException(
                "visual_media_asset_id_invalid",
                $"{path}.assetId must be the lowercase opaque asset_id returned by pptx_register_uploaded_image_asset.");
        }

        if (!MediaCropIntents.Contains(media.CropIntent))
        {
            throw new PptxValidationException(
                "visual_media_crop_invalid",
                $"{path}.cropIntent must be contain, cover, focalCenter, focalLeft, or focalRight.");
        }

        if (!MediaTextPositions.Contains(media.TextPosition))
        {
            throw new PptxValidationException(
                "visual_media_text_position_invalid",
                $"{path}.textPosition must be left or right.");
        }

        ValidateOptionalText(media.Caption, $"{path}.caption", 160);
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

    private static void ValidateVariantForSlide(VisualSlideSpec slide, string prefix)
    {
        var variant = slide.Variant;
        if (variant.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var supported = variant.ToLowerInvariant() switch
        {
            "split" => slide.Kind == VisualSlideKind.Media
                || slide.Kind == VisualSlideKind.Bullets
                && slide.Bullets?.Count >= 4
                && string.IsNullOrWhiteSpace(slide.Takeaway),
            "spotlight" => slide.Kind == VisualSlideKind.Metrics
                && slide.Metrics?.Count == 3
                || slide.Kind == VisualSlideKind.Cards
                && slide.Cards?.Count is >= 3 and <= 4,
            "editorial" => slide.Kind == VisualSlideKind.StructuredBrief
                && slide.Sections?.Count == 3,
            "stepped" => slide.Kind is VisualSlideKind.Timeline or VisualSlideKind.Roadmap
                && slide.Steps?.Count is >= 3 and <= 6,
            "pyramid" => slide.Kind == VisualSlideKind.Funnel
                && slide.Steps?.Count is >= 3 and <= 6,
            "loop" => slide.Kind == VisualSlideKind.Process
                && slide.Steps?.Count is >= 3 and <= 6,
            _ => false,
        };
        if (!supported)
        {
            throw new PptxValidationException(
                "visual_slide_variant_kind_mismatch",
                $"{prefix}.variant={variant} is not implemented for this slide's kind, content count, or takeaway state; use auto or satisfy the documented variant conditions.");
        }
    }

    private static void ValidateVisualObjectReferences(
        IReadOnlyList<VisualObjectAssetReference>? visualObjects,
        string prefix)
    {
        if (visualObjects is null)
        {
            return;
        }

        if (visualObjects.Count is < 1 or > 3)
        {
            throw new PptxValidationException(
                "visual_object_reference_count_invalid",
                $"{prefix}.visualObjects must contain between 1 and 3 prepared object references.");
        }

        if (visualObjects.Any(static item => item is null || !ImageAssetIdRegex().IsMatch(item.AssetId))
            || visualObjects.Select(static item => item.AssetId).Distinct(StringComparer.Ordinal).Count() != visualObjects.Count)
        {
            throw new PptxValidationException(
                "visual_object_reference_invalid",
                $"{prefix}.visualObjects must contain unique lowercase opaque asset IDs returned by pptx_prepare_visual_objects.");
        }
    }

    private static void ValidateDiagram(VisualDiagramSpec? diagram, string prefix)
    {
        if (diagram is null)
        {
            return;
        }

        var path = $"{prefix}.diagram";
        if (diagram.Nodes is null || diagram.Nodes.Count is < 2 or > 12)
        {
            throw new PptxValidationException(
                "visual_diagram_nodes_out_of_range",
                $"{path}.nodes must contain between 2 and 12 nodes.");
        }

        var countValid = diagram.Kind switch
        {
            VisualDiagramKind.Cycle => diagram.Nodes.Count is >= 3 and <= 6,
            VisualDiagramKind.Concentric => diagram.Nodes.Count is >= 2 and <= 4,
            VisualDiagramKind.Network => diagram.Nodes.Count is >= 3 and <= 9,
            VisualDiagramKind.Tree or VisualDiagramKind.Flow => diagram.Nodes.Count is >= 3 and <= 12,
            _ => false,
        };
        if (!countValid)
        {
            throw new PptxValidationException(
                "visual_diagram_kind_density_invalid",
                $"{path}.nodes count is unsafe for {diagram.Kind}; use cycle 3-6, concentric 2-4, network 3-9, or tree/flow 3-12.");
        }

        var directionValid = diagram.Kind == VisualDiagramKind.Cycle
            ? diagram.Direction.Equals("clockwise", StringComparison.OrdinalIgnoreCase)
            : diagram.Direction.Equals("leftToRight", StringComparison.OrdinalIgnoreCase)
                || diagram.Direction.Equals("topToBottom", StringComparison.OrdinalIgnoreCase);
        if (!directionValid)
        {
            throw new PptxValidationException(
                "visual_diagram_direction_invalid",
                $"{path}.direction is not supported for {diagram.Kind}.");
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        var emphasized = 0;
        for (var index = 0; index < diagram.Nodes.Count; index++)
        {
            var node = diagram.Nodes[index];
            var nodePath = $"{path}.nodes[{index}]";
            if (node is null || !OpaqueIdentifierRegex().IsMatch(node.Id) || !ids.Add(node.Id))
            {
                throw new PptxValidationException(
                    "visual_diagram_node_id_invalid",
                    $"{nodePath}.id must be unique and contain only ASCII letters, digits, hyphens, or underscores.");
            }

            ValidateText(node.Label, $"{nodePath}.label", 1, 48);
            ValidateOptionalText(node.Description, $"{nodePath}.description", 100);
            if (!IsSupportedTone(node.Tone))
            {
                throw new PptxValidationException(
                    "visual_diagram_node_tone_invalid",
                    $"{nodePath}.tone must be a supported semantic tone or a #RRGGBB color.");
            }

            emphasized += node.Emphasize ? 1 : 0;
        }

        if (emphasized > 2)
        {
            throw new PptxValidationException(
                "visual_diagram_emphasis_invalid",
                $"{path} may emphasize at most two nodes.");
        }

        var edges = diagram.Edges ?? [];
        if (edges.Count > 18)
        {
            throw new PptxValidationException(
                "visual_diagram_edges_out_of_range",
                $"{path}.edges must not contain more than 18 relationships.");
        }

        if (diagram.Kind is VisualDiagramKind.Tree or VisualDiagramKind.Flow && edges.Count < diagram.Nodes.Count - 1)
        {
            throw new PptxValidationException(
                "visual_diagram_edges_missing",
                $"{path}.edges must connect the tree or flow nodes.");
        }

        var edgeKeys = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < edges.Count; index++)
        {
            var edge = edges[index];
            var edgePath = $"{path}.edges[{index}]";
            if (edge is null
                || !ids.Contains(edge.From)
                || !ids.Contains(edge.To)
                || edge.From == edge.To
                || !edgeKeys.Add($"{edge.From}\0{edge.To}"))
            {
                throw new PptxValidationException(
                    "visual_diagram_edge_invalid",
                    $"{edgePath} must reference two different known node IDs and must not duplicate a relationship.");
            }

            ValidateOptionalText(edge.Label, $"{edgePath}.label", 32);
        }


        if (diagram.Kind is VisualDiagramKind.Tree or VisualDiagramKind.Flow)
        {
            var indegrees = ids.ToDictionary(static id => id, static _ => 0, StringComparer.Ordinal);
            foreach (var edge in edges)
            {
                indegrees[edge.To]++;
            }

            var queue = new Queue<string>(indegrees.Where(static pair => pair.Value == 0).Select(static pair => pair.Key));
            var visited = 0;
            while (queue.TryDequeue(out var nodeId))
            {
                visited++;
                foreach (var edge in edges.Where(edge => edge.From == nodeId))
                {
                    indegrees[edge.To]--;
                    if (indegrees[edge.To] == 0)
                    {
                        queue.Enqueue(edge.To);
                    }
                }
            }

            if (visited != ids.Count)
            {
                throw new PptxValidationException(
                    "visual_diagram_cycle_invalid",
                    $"{path}.edges for tree/flow must be acyclic. Use diagram.kind=cycle for a loop.");
            }
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

    private static void ValidateDataTable(
        VisualDataTableSpec? dataTable,
        string prefix,
        string density,
        bool hasTakeaway)
    {
        if (dataTable is null)
        {
            return;
        }

        var path = $"{prefix}.dataTable";
        var normalizedDensity = density.ToLowerInvariant();
        var maximumColumns = normalizedDensity switch
        {
            "airy" => 4,
            "detailed" => 6,
            _ => 5,
        };
        var maximumRows = normalizedDensity switch
        {
            "airy" => 6,
            "detailed" => 10,
            _ => 8,
        };
        var maximumCharacters = normalizedDensity switch
        {
            "airy" => 700,
            "detailed" => 1_600,
            _ => 1_100,
        };

        if (dataTable.Columns is null || dataTable.Columns.Count is < 2 || dataTable.Columns.Count > maximumColumns)
        {
            throw new PptxValidationException(
                "visual_data_table_columns_out_of_range",
                $"{path}.columns must contain between 2 and {maximumColumns} columns for {normalizedDensity} density.");
        }

        if (dataTable.Rows is null || dataTable.Rows.Count is < 1 || dataTable.Rows.Count > maximumRows)
        {
            throw new PptxValidationException(
                "visual_data_table_rows_out_of_range",
                $"{path}.rows must contain between 1 and {maximumRows} rows for {normalizedDensity} density.");
        }

        var totalCharacters = 0;
        for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
        {
            var column = dataTable.Columns[columnIndex];
            var columnPath = $"{path}.columns[{columnIndex}]";
            if (column is null)
            {
                throw new PptxValidationException(
                    "visual_data_table_column_invalid",
                    $"{columnPath} must not be null.");
            }

            ValidateSingleLineText(column.Header, $"{columnPath}.header", 1, 64);
            totalCharacters += column.Header.Length;
            if (!TableAlignments.Contains(column.Align))
            {
                throw new PptxValidationException(
                    "visual_data_table_alignment_invalid",
                    $"{columnPath}.align must be left, center, or right.");
            }

            if (!double.IsFinite(column.WidthWeight) || column.WidthWeight is < 0.5 or > 4)
            {
                throw new PptxValidationException(
                    "visual_data_table_width_invalid",
                $"{columnPath}.widthWeight must be between 0.5 and 4.0.");
            }
        }

        var geometry = GetDataTableGeometry(normalizedDensity, dataTable.Rows.Count, hasTakeaway);
        var totalWeight = dataTable.Columns.Sum(static column => column.WidthWeight);
        var columnWidths = dataTable.Columns
            .Select(column => geometry.TableWidth * column.WidthWeight / totalWeight)
            .ToArray();
        var minimumColumnWidth = normalizedDensity == "detailed" ? 0.85 : 1.0;
        for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
        {
            var effectiveWidth = columnWidths[columnIndex];
            var columnPath = $"{path}.columns[{columnIndex}]";
            if (effectiveWidth < minimumColumnWidth)
            {
                throw new PptxValidationException(
                    "visual_data_table_column_too_narrow",
                    $"{columnPath}.widthWeight produces a {effectiveWidth:F2}-inch column; rebalance weights so every {normalizedDensity} column is at least {minimumColumnWidth:F2} inches wide.");
            }

            var maximumHeaderCharacters = EstimateTableCellCharacterCapacity(
                effectiveWidth,
                geometry.HeaderHeight,
                geometry.FontSize,
                geometry.Margin,
                64);
            if (dataTable.Columns[columnIndex].Header.Length > maximumHeaderCharacters)
            {
                throw new PptxValidationException(
                    "visual_data_table_header_overflow_risk",
                    $"{columnPath}.header exceeds the safe capacity of {maximumHeaderCharacters} characters for its effective width; shorten the heading or widen the column.");
            }
        }

        for (var rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
        {
            var row = dataTable.Rows[rowIndex];
            var rowPath = $"{path}.rows[{rowIndex}]";
            if (row is null)
            {
                throw new PptxValidationException(
                    "visual_data_table_row_invalid",
                    $"{rowPath} must not be null.");
            }

            if (row.Cells is null || row.Cells.Count != dataTable.Columns.Count)
            {
                throw new PptxValidationException(
                    "visual_data_table_cells_mismatch",
                    $"{rowPath}.cells must contain exactly one cell per dataTable column.");
            }

            for (var cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                var cell = row.Cells[cellIndex];
                var cellPath = $"{rowPath}.cells[{cellIndex}]";
                if (cell is null)
                {
                    throw new PptxValidationException(
                        "visual_data_table_cell_invalid",
                        $"{cellPath} must not be null.");
                }

                var maximumCellCharacters = EstimateTableCellCharacterCapacity(
                    columnWidths[cellIndex],
                    geometry.RowHeight,
                    geometry.FontSize,
                    geometry.Margin,
                    100);
                ValidateSingleLineText(cell.Text, $"{cellPath}.text", 1, 100);
                if (cell.Text.Length > maximumCellCharacters)
                {
                    throw new PptxValidationException(
                        "visual_data_table_cell_overflow_risk",
                        $"{cellPath}.text exceeds the safe capacity of {maximumCellCharacters} characters for its rendered row and column; shorten it, widen the column, reduce rows, or split the table.");
                }
                totalCharacters += cell.Text.Length;
                if (!IsSupportedTone(cell.Tone))
                {
                    throw new PptxValidationException(
                        "visual_data_table_tone_invalid",
                        $"{cellPath}.tone must be a supported semantic tone or a #RRGGBB color.");
                }
            }
        }

        if (totalCharacters > maximumCharacters)
        {
            throw new PptxValidationException(
                "visual_content_density_invalid",
                $"{path} must not exceed {maximumCharacters} total characters for {normalizedDensity} density; split the table across slides.");
        }
    }

    private static DataTableGeometry GetDataTableGeometry(
        string density,
        int rowCount,
        bool hasTakeaway)
    {
        var outerX = density switch
        {
            "airy" => 0.82,
            "detailed" => 0.52,
            _ => 0.7,
        };
        var contentTop = density switch
        {
            "airy" => 2.12,
            "detailed" => 1.62,
            _ => 2.02,
        };
        var contentBottom = density switch
        {
            "airy" => 6.5,
            "detailed" => 6.9,
            _ => 6.62,
        };
        var tableBottom = hasTakeaway ? 6.58 : contentBottom;
        var headerHeight = density == "detailed" ? 0.52 : 0.62;
        var fontSize = density == "detailed" ? 9.2 : rowCount >= 7 ? 9.4 : 10.5;
        var margin = density == "detailed" ? 0.07 : 0.1;
        var rowHeight = (tableBottom - contentTop - headerHeight) / rowCount;
        return new DataTableGeometry(13.333 - outerX * 2, headerHeight, rowHeight, fontSize, margin);
    }

    private static int EstimateTableCellCharacterCapacity(
        double width,
        double height,
        double fontSize,
        double margin,
        int hardMaximum)
    {
        var lineHeight = fontSize / 72 * 1.2;
        var availableHeight = Math.Max(lineHeight, height - margin * 2);
        var maximumLines = Math.Max(1, (int)Math.Floor(availableHeight / lineHeight));
        var averageCharacterWidth = fontSize / 72 * 0.95;
        var charactersPerLine = Math.Max(4, (int)Math.Floor((width - margin * 2) / averageCharacterWidth));
        return Math.Clamp(charactersPerLine * maximumLines, 4, hardMaximum);
    }

    private readonly record struct DataTableGeometry(
        double TableWidth,
        double HeaderHeight,
        double RowHeight,
        double FontSize,
        double Margin);

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

    private static void ValidateOptionalContrast(
        string? foreground,
        string? background,
        double minimumRatio,
        string foregroundPath,
        string backgroundPath)
    {
        if (foreground is null || background is null)
        {
            return;
        }

        var foregroundLuminance = RelativeLuminance(foreground);
        var backgroundLuminance = RelativeLuminance(background);
        var contrastRatio = (Math.Max(foregroundLuminance, backgroundLuminance) + 0.05)
            / (Math.Min(foregroundLuminance, backgroundLuminance) + 0.05);
        if (contrastRatio < minimumRatio)
        {
            throw new PptxValidationException(
                "visual_theme_contrast_invalid",
                $"{foregroundPath} must have at least {minimumRatio:F1}:1 contrast against {backgroundPath}.");
        }
    }

    private static double RelativeLuminance(string color)
    {
        var normalized = color.TrimStart('#');
        static double Linearize(int component)
        {
            var value = component / 255d;
            return value <= 0.04045
                ? value / 12.92
                : Math.Pow((value + 0.055) / 1.055, 2.4);
        }

        var red = Linearize(Convert.ToInt32(normalized[..2], 16));
        var green = Linearize(Convert.ToInt32(normalized[2..4], 16));
        var blue = Linearize(Convert.ToInt32(normalized[4..6], 16));
        return 0.2126 * red + 0.7152 * green + 0.0722 * blue;
    }

    private static void ValidateOptionalText(string? value, string path, int maximumLength)
    {
        if (value is not null)
        {
            ValidateText(value, path, 0, maximumLength);
        }
    }

    private static void ValidateSpeakerNotes(VisualSpeakerNotesSpec? speakerNotes, string prefix)
    {
        if (speakerNotes is null)
        {
            return;
        }

        ValidateSingleLineText(speakerNotes.Purpose, $"{prefix}.speakerNotes.purpose", 1, 240);
        ValidateText(speakerNotes.TalkScript, $"{prefix}.speakerNotes.talkScript", 1, 1_200);
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

    private static void ValidateSingleLineText(string? value, string path, int minimumLength, int maximumLength)
    {
        ValidateText(value, path, minimumLength, maximumLength);
        if (value!.Contains('\n') || value.Contains('\r'))
        {
            throw new PptxValidationException(
                "visual_text_invalid",
                $"{path} must be a single line with no explicit line breaks.");
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

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex OpaqueIdentifierRegex();

    [GeneratedRegex("\\A[0-9a-f]{32}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex ImageAssetIdRegex();

    [GeneratedRegex("\\A[a-z][a-z0-9_-]{0,47}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex LocalVisualIdRegex();

    [GeneratedRegex("^(?:[1-9][0-9]?)/(?:1|2|4|8|16)$", RegexOptions.CultureInvariant)]
    private static partial Regex MusicTimeSignatureRegex();

    [GeneratedRegex("^([A-Ga-g])([#b]?)([0-8])$", RegexOptions.CultureInvariant)]
    private static partial Regex MusicPitchRegex();
}
