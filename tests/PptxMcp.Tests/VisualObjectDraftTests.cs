using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class VisualObjectDraftTests
{
    [Fact]
    public void BrandBoundDraftMaterializesPlannedObjectReferencesWhenTheSlideOmitsThem()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        try
        {
            var caller = new CallerContext("user-a", "conversation-a", null);
            var repository = new VisualObjectAssetRepository(TimeProvider.System);
            var assetId = Assert.Single(repository.Prepare(caller,
            [
                new VisualObjectBrief(
                    2,
                    VisualObjectPurpose.Emphasis,
                    VisualObjectArchetype.Frame,
                    PlacementRole: VisualObjectPlacementRole.FocusFrame),
            ]).Objects).AssetId;
            var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true);
            var catalog = new BrandProfileCatalog(options);
            var briefService = new DesignBriefService(options, TimeProvider.System, catalog);
            var summary = catalog.Query("general").Profiles.Single().Summary;
            var (brief, plan) = BrandProfileTestFactory.CreateBrief(
                new BrandProfileReference(summary.Id, summary.Version, summary.ContentHash),
                [assetId]);
            var validated = briefService.Validate(caller, brief, plan);
            var binding = briefService.AuthorizeForStart(caller, validated.BriefId, 2, "none", null, null)!;
            var drafts = new VisualDeckDraftService(
                Options.Create(new PptxMcpOptions { MaxSlides = 50 }),
                TimeProvider.System,
                visualObjectAssets: repository);
            var draft = drafts.Begin(
                caller,
                "Deck",
                2,
                binding.Theme,
                null,
                "ja-JP",
                binding.Design,
                "none",
                designBrief: binding);

            drafts.AddSlides(caller, draft.DraftId, null,
            [
                new VisualSlideSpec(
                    VisualSlideKind.Title,
                    "Decision",
                    Density: "airy",
                    RecipeId: "cover-airy"),
                new VisualSlideSpec(
                    VisualSlideKind.Metrics,
                    "Thresholds",
                    Metrics:
                    [
                        new VisualMetricSpec("42", "First"),
                        new VisualMetricSpec("7", "Second"),
                        new VisualMetricSpec("3", "Third"),
                    ],
                    Variant: "spotlight",
                    RecipeId: "kpi-balanced"),
            ]);

            var submission = drafts.AcquireForSubmission(caller, draft.DraftId);
            var materialized = Assert.Single(submission.Deck!.Slides[1].VisualObjects!);
            Assert.Equal(assetId, materialized.AssetId);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void DraftBindsOwnedObjectToPlannedSlideAndPersistsSemanticSnapshot()
    {
        var caller = new CallerContext("user-a", "conversation-a", null);
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var prepared = repository.Prepare(caller,
        [
            new VisualObjectBrief(
                1,
                VisualObjectPurpose.Emphasis,
                VisualObjectArchetype.Frame,
                PlacementRole: VisualObjectPlacementRole.FocusFrame),
        ]);
        var assetId = Assert.Single(prepared.Objects).AssetId;
        var drafts = new VisualDeckDraftService(
            Options.Create(new PptxMcpOptions { MaxSlides = 50 }),
            TimeProvider.System,
            visualObjectAssets: repository);
        var draft = drafts.Begin(caller, "Deck", 1, null, null, "ja-JP", null, "none");

        drafts.AddSlides(
            caller,
            draft.DraftId,
            null,
            [
                new VisualSlideSpec(
                    VisualSlideKind.Title,
                    "Title",
                    VisualObjects: [new VisualObjectAssetReference(assetId)]),
            ]);
        var submission = drafts.AcquireForSubmission(caller, draft.DraftId);

        Assert.NotNull(submission.Deck);
        var deck = submission.Deck!;
        Assert.NotNull(deck.VisualObjectAssets);
        var snapshot = Assert.Single(deck.VisualObjectAssets!);
        Assert.Equal(assetId, snapshot.AssetId);
        Assert.Equal(VisualObjectArchetype.Frame, snapshot.Brief.Archetype);
        Assert.Equal(64, snapshot.Fingerprint.Length);
    }

    [Fact]
    public void DraftRejectsObjectPreparedForAnotherSlide()
    {
        var caller = new CallerContext("user-a", "conversation-a", null);
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var assetId = Assert.Single(repository.Prepare(caller,
        [
            new VisualObjectBrief(2, VisualObjectPurpose.Direction, VisualObjectArchetype.Arrow),
        ]).Objects).AssetId;
        var drafts = new VisualDeckDraftService(
            Options.Create(new PptxMcpOptions { MaxSlides = 50 }),
            TimeProvider.System,
            visualObjectAssets: repository);
        var draft = drafts.Begin(caller, "Deck", 1, null, null, "ja-JP", null, "none");

        var error = Assert.Throws<PptxValidationException>(() => drafts.AddSlides(
            caller,
            draft.DraftId,
            null,
            [new VisualSlideSpec(VisualSlideKind.Title, "Title", VisualObjects: [new VisualObjectAssetReference(assetId)])]));

        Assert.Equal("visual_object_slide_mismatch", error.Code);
    }

    [Theory]
    [InlineData(VisualSlideKind.Bullets, VisualChartKind.Line, 2, 1)]
    [InlineData(VisualSlideKind.Chart, VisualChartKind.Bar, 2, 1)]
    [InlineData(VisualSlideKind.Chart, VisualChartKind.Line, 4, 1)]
    [InlineData(VisualSlideKind.Chart, VisualChartKind.Line, 2, 2)]
    public void DraftRejectsAnnotationPinOutsideAnExistingLineChartPoint(
        VisualSlideKind slideKind,
        VisualChartKind chartKind,
        int categoryOrdinal,
        int seriesOrdinal)
    {
        var caller = new CallerContext("user-a", "conversation-a", null);
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var assetId = Assert.Single(repository.Prepare(caller,
        [
            new VisualObjectBrief(
                1,
                VisualObjectPurpose.Annotation,
                VisualObjectArchetype.Callout,
                PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                Label: "交差点",
                Recipe: VisualObjectRecipe.AnnotationPin,
                AnchorCategoryOrdinal: categoryOrdinal,
                AnchorSeriesOrdinal: seriesOrdinal),
        ]).Objects).AssetId;
        var drafts = new VisualDeckDraftService(
            Options.Create(new PptxMcpOptions { MaxSlides = 50 }),
            TimeProvider.System,
            visualObjectAssets: repository);
        var draft = drafts.Begin(caller, "Deck", 1, null, null, "ja-JP", null, "none");
        var chart = new VisualChartSpec(
            chartKind,
            ["Q1", "Q2", "Q3"],
            [new VisualChartSeriesSpec("Series", [10, 20, 30])]);

        var error = Assert.Throws<PptxValidationException>(() => drafts.AddSlides(
            caller,
            draft.DraftId,
            null,
            [
                new VisualSlideSpec(
                    slideKind,
                    "Title",
                    Bullets: slideKind == VisualSlideKind.Bullets ? ["A"] : null,
                    Chart: slideKind == VisualSlideKind.Chart ? chart : null,
                    VisualObjects: [new VisualObjectAssetReference(assetId)]),
            ]));

        Assert.Equal("visual_object_chart_anchor_invalid", error.Code);
    }

    [Fact]
    public void DraftAcceptsAnnotationPinForAnExistingLineChartPoint()
    {
        var caller = new CallerContext("user-a", "conversation-a", null);
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var assetId = Assert.Single(repository.Prepare(caller,
        [
            new VisualObjectBrief(
                1,
                VisualObjectPurpose.Annotation,
                VisualObjectArchetype.Callout,
                PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                Label: "交差点",
                Recipe: VisualObjectRecipe.AnnotationPin,
                AnchorCategoryOrdinal: 2,
                AnchorSeriesOrdinal: 1),
        ]).Objects).AssetId;
        var drafts = new VisualDeckDraftService(
            Options.Create(new PptxMcpOptions { MaxSlides = 50 }),
            TimeProvider.System,
            visualObjectAssets: repository);
        var draft = drafts.Begin(caller, "Deck", 1, null, null, "ja-JP", null, "none");

        var result = drafts.AddSlides(caller, draft.DraftId, null,
        [
            new VisualSlideSpec(
                VisualSlideKind.Chart,
                "Title",
                Chart: new VisualChartSpec(
                    VisualChartKind.Line,
                    ["Q1", "Q2", "Q3"],
                    [new VisualChartSeriesSpec("Series", [10, 20, 30])]),
                VisualObjects: [new VisualObjectAssetReference(assetId)]),
        ]);

        Assert.Equal(1, result.AcceptedSlideCount);
    }
}
