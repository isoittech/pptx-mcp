using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class BrandProfileVisualDraftTests
{
    private static readonly CallerContext Caller = new("user-a", "conversation-a", "message-a");

    [Fact]
    public void EnforcesPlannedRecipeKindDensityAndVariantWhileAddingSlides()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true);
            var catalog = new BrandProfileCatalog(options);
            var briefs = new DesignBriefService(options, TimeProvider.System, catalog);
            var summary = catalog.Query("general").Profiles.Single().Summary;
            var (brief, plan) = BrandProfileTestFactory.CreateBrief(
                new BrandProfileReference(summary.Id, summary.Version, summary.ContentHash));
            var validated = briefs.Validate(Caller, brief, plan);
            var binding = briefs.AuthorizeForStart(Caller, validated.BriefId, 2, "none", null, null)!;
            var drafts = new VisualDeckDraftService(
                Options.Create(new PptxMcpOptions { MaxSlides = 50 }),
                TimeProvider.System);
            var started = drafts.Begin(
                Caller,
                "Recipe validation",
                2,
                binding.Theme,
                null,
                "en-US",
                binding.Design,
                "none",
                "auto",
                designBrief: binding);
            var densityMismatch = new VisualSlideSpec(
                VisualSlideKind.Title,
                "Decision",
                RecipeId: "cover-airy");

            var densityError = Assert.Throws<PptxValidationException>(() =>
                drafts.AddSlides(Caller, started.DraftId, null, [densityMismatch]));

            Assert.Equal("visual_slide_recipe_density_mismatch", densityError.Code);

            drafts.AddSlides(
                Caller,
                started.DraftId,
                null,
                [new VisualSlideSpec(
                    VisualSlideKind.Title,
                    "Decision",
                    Density: "airy",
                    RecipeId: "cover-airy")]);
            var kindMismatch = new VisualSlideSpec(
                VisualSlideKind.Cards,
                "Key values",
                Cards:
                [
                    new VisualCardSpec("One"),
                    new VisualCardSpec("Two"),
                    new VisualCardSpec("Three"),
                ],
                Variant: "spotlight",
                RecipeId: "kpi-balanced");
            var kindError = Assert.Throws<PptxValidationException>(() =>
                drafts.AddSlides(Caller, started.DraftId, null, [kindMismatch]));

            Assert.Equal("visual_slide_recipe_kind_mismatch", kindError.Code);

            var completed = drafts.AddSlides(
                Caller,
                started.DraftId,
                null,
                [new VisualSlideSpec(
                    VisualSlideKind.Metrics,
                    "Key values",
                    Metrics:
                    [
                        new VisualMetricSpec("42", "First"),
                        new VisualMetricSpec("7", "Second"),
                        new VisualMetricSpec("3", "Third"),
                    ],
                    Variant: "spotlight",
                    RecipeId: "kpi-balanced")]);

            Assert.Equal(0, completed.RemainingSlideCount);
            var submission = drafts.AcquireForSubmission(Caller, started.DraftId);
            Assert.NotNull(submission.Deck?.BrandProfileBinding);
            Assert.Equal(summary.ContentHash, submission.Deck.BrandProfileBinding.Profile.ContentHash);
            Assert.Equal(brief.SourcePolicy, submission.Deck.BrandProfileBinding.DesignBriefAudit?.SourcePolicy);
            Assert.Equal(brief.Assumptions, submission.Deck.BrandProfileBinding.DesignBriefAudit?.Assumptions);
            Assert.Equal(2, submission.Deck.BrandProfileBinding.DesignBriefAudit?.Slides.Count);
            Assert.Equal(
                ["cover-airy", "kpi-balanced"],
                submission.Deck.BrandProfileBinding.Slides.Select(static item => item.RecipeId));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }
}
