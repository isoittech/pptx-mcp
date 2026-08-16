using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Design;
using PptxMcp.Domain;

namespace PptxMcp.Tests;

internal static class BrandProfileTestFactory
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);

    public static string CreateRoot()
    {
        var root = Path.Combine(Path.GetTempPath(), $"pptx-mcp-brand-profiles-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        return root;
    }

    public static void WriteProfile(
        string root,
        string name = "General presentation profile",
        string description = "A generic profile for concise business presentations.",
        IReadOnlyList<string>? kpiRequiredAssetRoles = null,
        IReadOnlyList<string>? supportedDensities = null,
        int kpiItemCount = 3)
    {
        var bundle = Path.Combine(root, "general");
        Directory.CreateDirectory(bundle);
        var manifest = new
        {
            schema_version = 1,
            id = "general",
            version = "1.0.0",
            name,
            description,
            template_source = "none",
            template_id = "",
            color_roles = new
            {
                primary = "#17324D",
                secondary = "#49657D",
                accent = "#D96C36",
                background = "#F7F8FA",
                surface = "#FFFFFF",
                text = "#17212B",
                muted_text = "#66717C",
                positive = "#2B7A4B",
                warning = "#A35B00",
                critical = "#B42318",
                data_series = new[] { "#17324D", "#D96C36", "#49657D" },
            },
            typography = new
            {
                heading_font = "Aptos",
                body_font = "Aptos",
            },
            voice_rules = new[]
            {
                "Use concise conclusion-led headings.",
            },
            visual_rules = new
            {
                photography = Array.Empty<string>(),
                illustration = Array.Empty<string>(),
                iconography = new[] { "Use a restrained native icon set." },
                native_shapes = new[] { "Prefer editable diagrams for explanatory content." },
                tables = new[] { "Keep headers distinct and rows scannable." },
                charts = new[] { "Highlight only the decision-relevant series." },
                backgrounds = new[] { "Reserve dark backgrounds for section messages." },
                emphasis = new[] { "Use accent color sparingly." },
            },
            visual_object_policy = new
            {
                allowed_archetypes = new[] { "arrow", "curvedArrow", "frame", "callout", "bracket", "ring", "ribbon" },
                allowed_styles = new[] { "quietCorporate", "editorial", "technical" },
                default_style = "quietCorporate",
                maximum_per_slide = 3,
                maximum_per_deck = 16,
                strong_requires_focal_purpose = true,
            },
            prohibited_rules = new[]
            {
                "Do not use decorative placeholders as finished visuals.",
            },
            requires_confirmation_rules = new[]
            {
                "Confirm the use of identifiable people.",
            },
            approved_asset_collection_ids = new[]
            {
                "approved-general",
            },
            style_directions = new[]
            {
                new
                {
                    id = "standard",
                    name = "Standard",
                    summary = "Conclusion-led business slides with restrained visual emphasis.",
                    recommended_for = new[] { "cover", "kpi" },
                    design_style = "executive",
                    default_density = "balanced",
                    supported_densities = supportedDensities ?? ["airy", "balanced", "detailed"],
                    motif = "geometric",
                    theme_preset = "minimal",
                },
            },
            layout_recipes = new object[]
            {
                new
                {
                    id = "cover-airy",
                    purpose = "cover",
                    semantic_kind = "Title",
                    variant = "auto",
                    density = "airy",
                    style_direction_id = "standard",
                    required_asset_roles = Array.Empty<string>(),
                    sample_ids = new[] { "sample-cover-low" },
                },
                new
                {
                    id = "kpi-balanced",
                    purpose = "kpi",
                    semantic_kind = "Metrics",
                    variant = "spotlight",
                    density = "balanced",
                    style_direction_id = "standard",
                    required_asset_roles = kpiRequiredAssetRoles ?? [],
                    sample_ids = new[] { "sample-kpi-medium" },
                    item_count = kpiItemCount,
                },
            },
            samples = new[]
            {
                new
                {
                    id = "sample-cover-low",
                    title = "Short cover",
                    summary = "Low-information cover emphasizing one message.",
                    purpose = "cover",
                    density = "airy",
                    style_direction_id = "standard",
                    recipe_id = "cover-airy",
                    information_level = "low",
                },
                new
                {
                    id = "sample-kpi-medium",
                    title = "KPI summary",
                    summary = "Medium-information KPI composition with three key values.",
                    purpose = "kpi",
                    density = "balanced",
                    style_direction_id = "standard",
                    recipe_id = "kpi-balanced",
                    information_level = "medium",
                },
            },
        };
        File.WriteAllText(
            Path.Combine(bundle, "brand-profile.json"),
            JsonSerializer.Serialize(manifest, SerializerOptions));
    }

    public static IOptions<PptxMcpOptions> CreateOptions(
        string root,
        bool requireDesignBrief = false,
        int lifetimeMinutes = 60) =>
        Options.Create(new PptxMcpOptions
        {
            BrandProfilesRoot = root,
            RequireDesignBrief = requireDesignBrief,
            DesignBriefLifetimeMinutes = lifetimeMinutes,
            MaxSlides = 50,
        });

    public static BrandProfileCatalog CreateCatalog(
        string root,
        bool requireDesignBrief = false) =>
        new(CreateOptions(root, requireDesignBrief));

    public static (DesignBriefSpec Brief, AssetPlanItem[] AssetPlan) CreateBrief(
        BrandProfileReference profileReference,
        IReadOnlyList<string>? visualObjectAssetIds = null)
    {
        var brief = new DesignBriefSpec(
            "Business leaders",
            "Reach a decision on the proposed operating plan.",
            DesignDeliveryMode.Projection,
            "Clear and evidence-led",
            "balanced",
            profileReference,
            "standard",
            "Lead with the conclusion, then use editable KPI objects as evidence.",
            DesignSourcePolicy.ApprovedOnly,
            2,
            [new DesignAssumption("The presentation will be projected in a meeting.", DesignAssumptionStatus.Inferred)],
            []);
        var assetPlan = new[]
        {
            new AssetPlanItem(
                1,
                "cover",
                "cover-airy",
                AssetVisualPurpose.None,
                AssetPreferredMedium.None,
                AssetAcquisition.None,
                AssetFallback.None,
                AssetPlanStatus.Omitted),
            new AssetPlanItem(
                2,
                "kpi",
                "kpi-balanced",
                AssetVisualPurpose.Data,
                AssetPreferredMedium.NativeDiagram,
                AssetAcquisition.NativeDraw,
                AssetFallback.None,
                AssetPlanStatus.Ready,
                VisualObjectAssetIds: visualObjectAssetIds),
        };
        return (brief, assetPlan);
    }
}
