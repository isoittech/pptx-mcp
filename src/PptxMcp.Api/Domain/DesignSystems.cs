using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PptxMcp.Domain;

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<DesignDeliveryMode>))]
public enum DesignDeliveryMode
{
    Projection,
    Handout,
    Both,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<DesignAssumptionStatus>))]
public enum DesignAssumptionStatus
{
    Confirmed,
    Inferred,
    NeedsConfirmation,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<DesignSourcePolicy>))]
public enum DesignSourcePolicy
{
    ApprovedOnly,
    ApprovedOrUserProvided,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<DesignBriefSelectionSource>))]
public enum DesignBriefSelectionSource
{
    AgentDefault,
    UserCard,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<AssetVisualPurpose>))]
public enum AssetVisualPurpose
{
    Evidence,
    Atmosphere,
    Comparison,
    Process,
    Relationship,
    Location,
    Person,
    Product,
    Data,
    None,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<AssetPreferredMedium>))]
public enum AssetPreferredMedium
{
    Photo,
    Illustration,
    Icon,
    NativeDiagram,
    Chart,
    None,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<AssetAcquisition>))]
public enum AssetAcquisition
{
    NativeDraw,
    UserUpload,
    ApprovedLibrary,
    None,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<AssetFallback>))]
public enum AssetFallback
{
    NativeDraw,
    NoAssetLayout,
    AskUser,
    None,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<AssetPlanStatus>))]
public enum AssetPlanStatus
{
    Ready,
    NeedsUser,
    FallbackSelected,
    Omitted,
}

[JsonConverter(typeof(CamelCaseJsonStringEnumConverter<AssetLicenseStatus>))]
public enum AssetLicenseStatus
{
    NotRequired,
    Approved,
    UserProvided,
    Unknown,
    Restricted,
}

public sealed record BrandProfileReference(
    [property: JsonPropertyName("id"), Description("Opaque Brand Profile ID copied from pptx_get_design_catalog.")]
    string Id,
    [property: JsonPropertyName("version"), Description("Immutable profile version copied from pptx_get_design_catalog.")]
    string Version,
    [property: JsonPropertyName("content_hash"), Description("Exact SHA-256 content hash copied from pptx_get_design_catalog.")]
    string ContentHash);

public sealed record DesignAssumption(
    [property: JsonPropertyName("text"), Description("A concise confirmed or inferred assumption. Never include URLs, file paths, credentials, or private endpoints.")]
    string Text,
    [property: JsonPropertyName("status"), Description("confirmed, inferred, or needsConfirmation. A finalized brief cannot contain needsConfirmation.")]
    DesignAssumptionStatus Status);

public sealed record DesignBriefSpec(
    [property: JsonPropertyName("audience"), Description("Primary audience and decision-making context.")]
    string Audience,
    [property: JsonPropertyName("purpose"), Description("Presentation objective and intended outcome.")]
    string Purpose,
    [property: JsonPropertyName("delivery_mode"), Description("projection, handout, or both.")]
    DesignDeliveryMode DeliveryMode,
    [property: JsonPropertyName("desired_tone"), Description("Concise desired writing and visual tone.")]
    string DesiredTone,
    [property: JsonPropertyName("density"), Description("Base information density: airy, balanced, or detailed.")]
    string Density,
    [property: JsonPropertyName("brand_profile"), Description("Exact immutable Brand Profile reference from the design catalog.")]
    BrandProfileReference BrandProfile,
    [property: JsonPropertyName("style_direction_id"), Description("Style direction ID from the selected Brand Profile.")]
    string StyleDirectionId,
    [property: JsonPropertyName("visual_strategy"), Description("Short explanation of how visuals, hierarchy, and composition support the story.")]
    string VisualStrategy,
    [property: JsonPropertyName("source_policy"), Description("approvedOnly or approvedOrUserProvided. Phase 1 does not fetch web or generated images.")]
    DesignSourcePolicy SourcePolicy,
    [property: JsonPropertyName("expected_slide_count"), Description("Exact final slide count that will be passed to pptx_start_visual_deck.")]
    int ExpectedSlideCount,
    [property: JsonPropertyName("assumptions"), Description("Confirmed and explicitly inferred assumptions. Do not mix items that still need user confirmation.")]
    IReadOnlyList<DesignAssumption> Assumptions,
    [property: JsonPropertyName("questions_for_user"), Description("Must be empty before validation. Ask only material questions, then submit the finalized brief without unresolved questions.")]
    IReadOnlyList<string> QuestionsForUser);

public sealed record AssetPlanItem(
    [property: JsonPropertyName("slide_number"), Description("One-based target slide number. The final plan contains exactly one item for every slide.")]
    int SlideNumber,
    [property: JsonPropertyName("purpose"), Description("Opaque layout-purpose ID that must match the chosen profile recipe purpose.")]
    string Purpose,
    [property: JsonPropertyName("recipe_id"), Description("Exact immutable layout recipe ID from the selected Brand Profile.")]
    string RecipeId,
    [property: JsonPropertyName("visual_purpose"), Description("evidence, atmosphere, comparison, process, relationship, location, person, product, data, or none.")]
    AssetVisualPurpose VisualPurpose,
    [property: JsonPropertyName("preferred_medium"), Description("photo, illustration, icon, nativeDiagram, chart, or none. acquisition=none requires preferred_medium=none.")]
    AssetPreferredMedium PreferredMedium,
    [property: JsonPropertyName("acquisition"), Description("Phase-1 planning choice: nativeDraw, userUpload, approvedLibrary, or none. Exact no-asset combination: acquisition=none, preferred_medium=none, fallback=none, status=omitted, license_status=notRequired. This server does not insert images yet.")]
    AssetAcquisition Acquisition,
    [property: JsonPropertyName("fallback"), Description("Safe fallback: nativeDraw, noAssetLayout, askUser, or none. IMPORTANT: acquisition=none requires fallback=none; noAssetLayout is only a fallbackSelected replacement for userUpload or approvedLibrary, never for acquisition=none.")]
    AssetFallback Fallback,
    [property: JsonPropertyName("status"), Description("ready, needsUser, fallbackSelected, or omitted. acquisition=none requires omitted; nativeDraw requires ready; userUpload or approvedLibrary requires fallbackSelected in phase 1. A finalized plan cannot remain needsUser.")]
    AssetPlanStatus Status,
    [property: JsonPropertyName("license_status"), Description("notRequired, approved, userProvided, unknown, or restricted. acquisition=none and nativeDraw require notRequired. Unknown or restricted assets must not be treated as usable.")]
    AssetLicenseStatus LicenseStatus = AssetLicenseStatus.NotRequired,
    [property: JsonPropertyName("approved_asset_collection_id"), Description("Optional opaque approved-library collection ID from the Brand Profile. Never provide a URL or path.")]
    string? ApprovedAssetCollectionId = null,
    [property: JsonPropertyName("attribution_ref"), Description("Optional opaque attribution record ID. Never provide source text, a URL, or a path here.")]
    string? AttributionRef = null,
    [property: JsonPropertyName("crop_intent"), Description("Optional crop token: contain, cover, focalCenter, focalLeft, focalRight, or none.")]
    string? CropIntent = null,
    [property: JsonPropertyName("aspect_ratio"), Description("Optional aspect-ratio token: landscape16x9, landscape4x3, square1x1, portrait4x5, or flexible.")]
    string? AspectRatio = null,
    [property: JsonPropertyName("text_safe_area"), Description("Optional text-safe-area token: none, left, right, top, or bottom.")]
    string? TextSafeArea = null);

public sealed record DesignBriefStyleAlternative(
    [property: JsonPropertyName("style_direction_id"), Description("Alternative style direction ID from the same immutable Brand Profile.")]
    string StyleDirectionId,
    [property: JsonPropertyName("density"), Description("Alternative base density: airy, balanced, or detailed.")]
    string Density,
    [property: JsonPropertyName("recipe_ids"), Description("Exactly one recipe ID per slide, in slide order. Only the recipe IDs differ from the common Asset Plan.")]
    IReadOnlyList<string> RecipeIds);

public sealed record DesignBriefNoPhotoOverride(
    [property: JsonPropertyName("slide_number"), Description("One-based slide number whose planned photo role is replaced by a no-image composition.")]
    int SlideNumber,
    [property: JsonPropertyName("recipe_id"), Description("No-image composition recipe ID for this slide from the recommended style direction.")]
    string RecipeId,
    [property: JsonPropertyName("preferred_medium"), Description("No-image composition medium: icon, nativeDiagram, chart, or none.")]
    AssetPreferredMedium PreferredMedium,
    [property: JsonPropertyName("acquisition"), Description("No-image composition acquisition: nativeDraw or none.")]
    AssetAcquisition Acquisition);

public sealed record DesignBriefValidationView(
    [property: JsonPropertyName("brief_id")] string BriefId,
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("expires_at")] DateTimeOffset ExpiresAt,
    [property: JsonPropertyName("brand_profile")] BrandProfileReference BrandProfile,
    [property: JsonPropertyName("style_direction_id")] string StyleDirectionId,
    [property: JsonPropertyName("expected_slide_count")] int ExpectedSlideCount,
    [property: JsonPropertyName("asset_plan_summary")] AssetPlanSummary AssetPlanSummary,
    [property: JsonPropertyName("instruction")] string Instruction);

public sealed record DesignBriefSelectionCancellationView(
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("instruction")] string Instruction);

public sealed record AssetPlanSummary(
    [property: JsonPropertyName("native_draw_count")] int NativeDrawCount,
    [property: JsonPropertyName("no_asset_count")] int NoAssetCount,
    [property: JsonPropertyName("fallback_selected_count")] int FallbackSelectedCount);

public sealed record BrandProfileCatalogSummary(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("version")] string Version,
    [property: JsonPropertyName("content_hash")] string ContentHash,
    [property: JsonPropertyName("description")] string Description,
    [property: JsonPropertyName("template_source")] string TemplateSource,
    [property: JsonPropertyName("style_direction_ids")] IReadOnlyList<string> StyleDirectionIds);

public sealed record BrandColorRoles(
    [property: JsonPropertyName("primary")] string Primary,
    [property: JsonPropertyName("secondary")] string Secondary,
    [property: JsonPropertyName("accent")] string Accent,
    [property: JsonPropertyName("background")] string Background,
    [property: JsonPropertyName("surface")] string Surface,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("muted_text")] string MutedText,
    [property: JsonPropertyName("positive")] string Positive,
    [property: JsonPropertyName("warning")] string Warning,
    [property: JsonPropertyName("critical")] string Critical,
    [property: JsonPropertyName("data_series")] IReadOnlyList<string> DataSeries);

public sealed record BrandTypography(
    [property: JsonPropertyName("heading_font")] string HeadingFont,
    [property: JsonPropertyName("body_font")] string BodyFont);

public sealed record BrandStyleDirection(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("summary")] string Summary,
    [property: JsonPropertyName("recommended_for")] IReadOnlyList<string> RecommendedFor,
    [property: JsonPropertyName("design_style")] string DesignStyle,
    [property: JsonPropertyName("default_density")] string DefaultDensity,
    [property: JsonPropertyName("supported_densities")] IReadOnlyList<string> SupportedDensities,
    [property: JsonPropertyName("motif")] string Motif,
    [property: JsonPropertyName("theme_preset")] string ThemePreset);

public sealed record BrandLayoutRecipe(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("purpose")] string Purpose,
    [property: JsonPropertyName("semantic_kind")] VisualSlideKind SemanticKind,
    [property: JsonPropertyName("variant")] string Variant,
    [property: JsonPropertyName("density")] string Density,
    [property: JsonPropertyName("style_direction_id")] string StyleDirectionId,
    [property: JsonPropertyName("required_asset_roles")] IReadOnlyList<string> RequiredAssetRoles,
    [property: JsonPropertyName("sample_ids")] IReadOnlyList<string> SampleIds,
    [property: JsonPropertyName("item_count"), Description("Optional exact repeated-item count required by the recipe. Metrics spotlight recipes must set this to 3.")]
    int? ItemCount = null);

public sealed record BrandSampleSummary(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("title")] string Title,
    [property: JsonPropertyName("summary")] string Summary,
    [property: JsonPropertyName("purpose")] string Purpose,
    [property: JsonPropertyName("density")] string Density,
    [property: JsonPropertyName("style_direction_id")] string StyleDirectionId,
    [property: JsonPropertyName("recipe_id")] string RecipeId,
    [property: JsonPropertyName("information_level")] string InformationLevel);

public sealed record BrandSampleThumbnail(
    string SampleId,
    string MimeType,
    int Width,
    int Height,
    [property: JsonIgnore] ReadOnlyMemory<byte> Bytes,
    [property: JsonIgnore] long DecodedBytes);

public sealed record BrandVisualRuleSet(
    [property: JsonPropertyName("photography")] IReadOnlyList<string> Photography,
    [property: JsonPropertyName("illustration")] IReadOnlyList<string> Illustration,
    [property: JsonPropertyName("iconography")] IReadOnlyList<string> Iconography,
    [property: JsonPropertyName("native_shapes")] IReadOnlyList<string> NativeShapes,
    [property: JsonPropertyName("tables")] IReadOnlyList<string> Tables,
    [property: JsonPropertyName("charts")] IReadOnlyList<string> Charts,
    [property: JsonPropertyName("backgrounds")] IReadOnlyList<string> Backgrounds,
    [property: JsonPropertyName("emphasis")] IReadOnlyList<string> Emphasis);

public sealed record BrandProfileCatalogDetail(
    [property: JsonPropertyName("summary")] BrandProfileCatalogSummary Summary,
    [property: JsonPropertyName("color_roles")] BrandColorRoles ColorRoles,
    [property: JsonPropertyName("typography")] BrandTypography Typography,
    [property: JsonPropertyName("voice_rules")] IReadOnlyList<string> VoiceRules,
    [property: JsonPropertyName("visual_rules")] BrandVisualRuleSet VisualRules,
    [property: JsonPropertyName("prohibited_rules")] IReadOnlyList<string> ProhibitedRules,
    [property: JsonPropertyName("requires_confirmation_rules")] IReadOnlyList<string> RequiresConfirmationRules,
    [property: JsonPropertyName("approved_asset_collection_ids")] IReadOnlyList<string> ApprovedAssetCollectionIds,
    [property: JsonPropertyName("style_directions")] IReadOnlyList<BrandStyleDirection> StyleDirections);

public sealed record DesignCatalogProfileView(
    [property: JsonPropertyName("summary")] BrandProfileCatalogSummary Summary,
    [property: JsonPropertyName("detail")] BrandProfileCatalogDetail? Detail,
    [property: JsonPropertyName("recipes")] IReadOnlyList<BrandLayoutRecipe> Recipes,
    [property: JsonPropertyName("samples")] IReadOnlyList<BrandSampleSummary> Samples);

public sealed record DesignCatalogView(
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("profiles")] IReadOnlyList<DesignCatalogProfileView> Profiles,
    [property: JsonPropertyName("instruction")] string Instruction);

public sealed record BrandProfileSnapshot(
    BrandProfileCatalogDetail Detail,
    string TemplateId,
    IReadOnlyList<BrandLayoutRecipe> LayoutRecipes,
    IReadOnlyList<BrandSampleSummary> Samples,
    [property: JsonIgnore] IReadOnlyDictionary<string, BrandSampleThumbnail> SampleThumbnails)
{
    public string Id => Detail.Summary.Id;

    public string Version => Detail.Summary.Version;

    public string ContentHash => Detail.Summary.ContentHash;

    public string TemplateSource => Detail.Summary.TemplateSource;
}

public sealed record ValidatedDesignBriefBinding(
    string BriefId,
    DesignBriefSpec Brief,
    BrandProfileSnapshot Profile,
    BrandStyleDirection StyleDirection,
    VisualThemeSpec Theme,
    VisualDesignSpec Design,
    IReadOnlyDictionary<int, AssetPlanItem> AssetPlan,
    DateTimeOffset ExpiresAt,
    DesignBriefSelectionSource SelectionSource = DesignBriefSelectionSource.AgentDefault);

public sealed record VisualDeckBrandProfileBinding(
    [property: JsonPropertyName("profile")] BrandProfileReference Profile,
    [property: JsonPropertyName("style_direction_id")] string StyleDirectionId,
    [property: JsonPropertyName("slides")] IReadOnlyList<VisualSlideRecipeBinding> Slides,
    [property: JsonPropertyName("design_brief_audit")] VisualDeckDesignBriefAudit? DesignBriefAudit = null);

public sealed record VisualSlideRecipeBinding(
    [property: JsonPropertyName("slide_number")] int SlideNumber,
    [property: JsonPropertyName("recipe_id")] string RecipeId,
    [property: JsonPropertyName("semantic_kind")] VisualSlideKind SemanticKind,
    [property: JsonPropertyName("density")] string Density,
    [property: JsonPropertyName("variant")] string Variant);

public sealed record VisualDeckDesignBriefAudit(
    [property: JsonPropertyName("source_policy")] DesignSourcePolicy SourcePolicy,
    [property: JsonPropertyName("assumptions")] IReadOnlyList<DesignAssumption> Assumptions,
    [property: JsonPropertyName("slides")] IReadOnlyList<VisualSlideAssetAudit> Slides,
    [property: JsonPropertyName("selection_source")] DesignBriefSelectionSource SelectionSource = DesignBriefSelectionSource.AgentDefault);

public sealed record VisualSlideAssetAudit(
    [property: JsonPropertyName("slide_number")] int SlideNumber,
    [property: JsonPropertyName("visual_purpose")] AssetVisualPurpose VisualPurpose,
    [property: JsonPropertyName("preferred_medium")] AssetPreferredMedium PreferredMedium,
    [property: JsonPropertyName("acquisition")] AssetAcquisition Acquisition,
    [property: JsonPropertyName("fallback")] AssetFallback Fallback,
    [property: JsonPropertyName("status")] AssetPlanStatus Status,
    [property: JsonPropertyName("license_status")] AssetLicenseStatus LicenseStatus,
    [property: JsonPropertyName("approved_asset_collection_id")] string? ApprovedAssetCollectionId,
    [property: JsonPropertyName("attribution_ref")] string? AttributionRef,
    [property: JsonPropertyName("crop_intent")] string? CropIntent,
    [property: JsonPropertyName("aspect_ratio")] string? AspectRatio,
    [property: JsonPropertyName("text_safe_area")] string? TextSafeArea);

public sealed class CamelCaseJsonStringEnumConverter<TEnum>()
    : JsonStringEnumConverter<TEnum>(JsonNamingPolicy.CamelCase, allowIntegerValues: false)
    where TEnum : struct, Enum;
