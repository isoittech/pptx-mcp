using System.Text.Json;
using System.Text.Json.Nodes;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class BrandProfileJobBindingTests
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);

    [Fact]
    public void ProfileBoundRefinementPreservesExactRecipeContract()
    {
        var original = CreateBoundDeck();
        var validRevision = new VisualSlideRevision(
            1,
            new VisualSlideSpec(
                VisualSlideKind.Title,
                "Updated decision",
                Density: "airy",
                RecipeId: "cover-airy"));

        var result = JobService.ApplyVisualDeckRevisions(original, [validRevision], 50);

        Assert.Equal("Updated decision", result.Slides[0].Title);
        Assert.Same(original.BrandProfileBinding, result.BrandProfileBinding);
    }

    [Fact]
    public void ProfileBoundRefinementRejectsRecipeKindDensityAndVariantDrift()
    {
        var original = CreateBoundDeck();
        var recipeError = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyVisualDeckRevisions(
                original,
                [new VisualSlideRevision(
                    1,
                    new VisualSlideSpec(
                        VisualSlideKind.Title,
                        "Changed recipe",
                        Density: "airy",
                        RecipeId: "different"))],
                50));
        var kindError = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyVisualDeckRevisions(
                original,
                [new VisualSlideRevision(
                    1,
                    new VisualSlideSpec(
                        VisualSlideKind.Section,
                        "Changed kind",
                        Density: "airy",
                        RecipeId: "cover-airy"))],
                50));
        var densityError = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyVisualDeckRevisions(
                original,
                [new VisualSlideRevision(
                    1,
                    new VisualSlideSpec(
                        VisualSlideKind.Title,
                        "Changed density",
                        Density: "detailed",
                        RecipeId: "cover-airy"))],
                50));
        var variantError = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyVisualDeckRevisions(
                original,
                [new VisualSlideRevision(
                    1,
                    new VisualSlideSpec(
                        VisualSlideKind.Title,
                        "Changed variant",
                        Density: "airy",
                        Variant: "editorial",
                        RecipeId: "cover-airy"))],
                50));

        Assert.Equal("visual_slide_recipe_mismatch", recipeError.Code);
        Assert.Equal("visual_slide_recipe_kind_mismatch", kindError.Code);
        Assert.Equal("visual_slide_recipe_density_mismatch", densityError.Code);
        Assert.Equal("visual_slide_recipe_variant_mismatch", variantError.Code);
    }

    [Fact]
    public void ProfileBoundDeckRejectsUnplannedInsertionInPhaseOne()
    {
        var original = CreateBoundDeck();

        var error = Assert.Throws<PptxValidationException>(() =>
            JobService.InsertVisualSlides(
                original,
                [new VisualSlideSpec(VisualSlideKind.Statement, "Unplanned", Body: "Not in Asset Plan")],
                null,
                50));

        Assert.Equal("brand_profile_insert_requires_design_brief", error.Code);
    }

    [Fact]
    public void AuditAndRecipeBindingSurviveJobPayloadRoundTripAndStillProtectRefinement()
    {
        var original = CreateBoundDeck();

        var json = JsonSerializer.Serialize(original, SerializerOptions);
        var restored = JsonSerializer.Deserialize<VisualDeckSpec>(json, SerializerOptions);

        Assert.NotNull(restored?.BrandProfileBinding?.DesignBriefAudit);
        Assert.DoesNotContain("brief_id", json, StringComparison.Ordinal);
        Assert.DoesNotContain("choiceSessionId", json, StringComparison.Ordinal);
        Assert.DoesNotContain("optionId", json, StringComparison.Ordinal);
        Assert.Equal(
            DesignBriefSelectionSource.AgentDefault,
            restored.BrandProfileBinding.DesignBriefAudit.SelectionSource);
        Assert.Equal(
            DesignSourcePolicy.ApprovedOnly,
            restored.BrandProfileBinding.DesignBriefAudit.SourcePolicy);
        Assert.Equal(
            DesignAssumptionStatus.Inferred,
            Assert.Single(restored.BrandProfileBinding.DesignBriefAudit.Assumptions).Status);
        Assert.Equal(2, restored.BrandProfileBinding.DesignBriefAudit.Slides.Count);
        var assetAudit = restored.BrandProfileBinding.DesignBriefAudit.Slides[0];
        Assert.Equal(AssetAcquisition.NativeDraw, assetAudit.Acquisition);
        Assert.Equal(AssetLicenseStatus.NotRequired, assetAudit.LicenseStatus);

        var error = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyVisualDeckRevisions(
                restored,
                [new VisualSlideRevision(
                    1,
                    new VisualSlideSpec(
                        VisualSlideKind.Title,
                        "Changed recipe",
                        Density: "airy",
                        RecipeId: "different"))],
                50));

        Assert.Equal("visual_slide_recipe_mismatch", error.Code);
    }

    [Fact]
    public void LegacyJobPayloadWithoutSelectionSourceDefaultsToAgentDefault()
    {
        var json = JsonNode.Parse(JsonSerializer.Serialize(CreateBoundDeck(), SerializerOptions))!.AsObject();
        json["brand_profile_binding"]!["design_brief_audit"]!
            .AsObject()
            .Remove("selection_source");

        var restored = JsonSerializer.Deserialize<VisualDeckSpec>(json.ToJsonString(), SerializerOptions);

        Assert.Equal(
            DesignBriefSelectionSource.AgentDefault,
            restored?.BrandProfileBinding?.DesignBriefAudit?.SelectionSource);
    }

    private static VisualDeckSpec CreateBoundDeck()
    {
        var profile = new BrandProfileReference("general", "1.0.0", new string('a', 64));
        var binding = new VisualDeckBrandProfileBinding(
            profile,
            "standard",
            [
                new VisualSlideRecipeBinding(1, "cover-airy", VisualSlideKind.Title, "airy", "auto"),
                new VisualSlideRecipeBinding(2, "kpi-balanced", VisualSlideKind.Metrics, "balanced", "spotlight"),
            ],
            new VisualDeckDesignBriefAudit(
                DesignSourcePolicy.ApprovedOnly,
                [new DesignAssumption("Projected in a meeting", DesignAssumptionStatus.Inferred)],
                [
                    new VisualSlideAssetAudit(
                        1,
                        AssetVisualPurpose.None,
                        AssetPreferredMedium.NativeDiagram,
                        AssetAcquisition.NativeDraw,
                        AssetFallback.None,
                        AssetPlanStatus.Ready,
                        AssetLicenseStatus.NotRequired,
                        null,
                        null,
                        "none",
                        "landscape16x9",
                        "none"),
                    new VisualSlideAssetAudit(
                        2,
                        AssetVisualPurpose.Data,
                        AssetPreferredMedium.Chart,
                        AssetAcquisition.NativeDraw,
                        AssetFallback.None,
                        AssetPlanStatus.Ready,
                        AssetLicenseStatus.NotRequired,
                        null,
                        null,
                        "none",
                        "landscape16x9",
                        "none"),
                ]));
        return new VisualDeckSpec(
            "Bound deck",
            [
                new VisualSlideSpec(
                    VisualSlideKind.Title,
                    "Decision",
                    Density: "airy",
                    RecipeId: "cover-airy"),
                new VisualSlideSpec(
                    VisualSlideKind.Metrics,
                    "Evidence",
                    Metrics:
                    [
                        new VisualMetricSpec("1", "First"),
                        new VisualMetricSpec("2", "Second"),
                        new VisualMetricSpec("3", "Third"),
                    ],
                    Variant: "spotlight",
                    RecipeId: "kpi-balanced"),
            ],
            Design: new VisualDesignSpec(Density: "balanced"),
            RendererContract: "visual-v5",
            BrandProfileBinding: binding);
    }
}
