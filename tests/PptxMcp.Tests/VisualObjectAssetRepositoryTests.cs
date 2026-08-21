using PptxMcp.Domain;
using PptxMcp.Storage;
using ModelContextProtocol.Protocol;
using PptxMcp.Design;

namespace PptxMcp.Tests;

public sealed class VisualObjectAssetRepositoryTests
{
    [Fact]
    public void PreparedObjectsAreBoundToUserConversationAndExpiry()
    {
        var clock = new TestTimeProvider(new DateTimeOffset(2026, 8, 16, 0, 0, 0, TimeSpan.Zero));
        var repository = new VisualObjectAssetRepository(clock);
        var owner = new CallerContext("user-a", "conversation-a", null);
        var prepared = repository.Prepare(owner,
        [
            new VisualObjectBrief(
                2,
                VisualObjectPurpose.Annotation,
                VisualObjectArchetype.Callout,
                PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                Label: "要確認"),
        ]);
        var assetId = Assert.Single(prepared.Objects).AssetId;

        Assert.Equal(assetId, repository.GetOwned(owner, assetId).AssetId);
        Assert.Equal(
            "visual_object_asset_not_found",
            Assert.Throws<PptxValidationException>(() => repository.GetOwned(
                new CallerContext("user-b", "conversation-a", null),
                assetId)).Code);
        Assert.Equal(
            "visual_object_asset_not_found",
            Assert.Throws<PptxValidationException>(() => repository.GetOwned(
                new CallerContext("user-a", "conversation-b", null),
                assetId)).Code);

        clock.Advance(TimeSpan.FromHours(2));
        Assert.Equal(
            "visual_object_asset_not_found",
            Assert.Throws<PptxValidationException>(() => repository.GetOwned(owner, assetId)).Code);
    }

    [Fact]
    public void BatchAndPerSlideLimitsFailBeforeCreatingAssets()
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var caller = new CallerContext("user-a", "conversation-a", null);
        var tooMany = Enumerable.Range(1, VisualObjectAssetRepository.MaximumBatchObjects + 1)
            .Select(index => new VisualObjectBrief(
                index,
                VisualObjectPurpose.Direction,
                VisualObjectArchetype.Arrow))
            .ToArray();
        Assert.Equal(
            "visual_object_batch_invalid",
            Assert.Throws<PptxValidationException>(() => repository.Prepare(caller, tooMany)).Code);

        var crowdedSlide = Enumerable.Range(0, VisualObjectAssetRepository.MaximumObjectsPerSlide + 1)
            .Select(_ => new VisualObjectBrief(
                1,
                VisualObjectPurpose.Emphasis,
                VisualObjectArchetype.Frame))
            .ToArray();
        Assert.Equal(
            "visual_object_slide_limit_invalid",
            Assert.Throws<PptxValidationException>(() => repository.Prepare(caller, crowdedSlide)).Code);
    }

    [Fact]
    public void SemanticMismatchAndMultipleStrongObjectsAreRejected()
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var caller = new CallerContext("user-a", "conversation-a", null);
        Assert.Equal(
            "visual_object_semantics_invalid",
            Assert.Throws<PptxValidationException>(() => repository.Prepare(caller,
            [
                new VisualObjectBrief(1, VisualObjectPurpose.Cycle, VisualObjectArchetype.Frame),
            ])).Code);
        Assert.Equal(
            "visual_object_emphasis_invalid",
            Assert.Throws<PptxValidationException>(() => repository.Prepare(caller,
            [
                new VisualObjectBrief(1, VisualObjectPurpose.Emphasis, VisualObjectArchetype.Frame, Emphasis: VisualObjectEmphasis.Strong),
                new VisualObjectBrief(1, VisualObjectPurpose.Annotation, VisualObjectArchetype.Callout, Emphasis: VisualObjectEmphasis.Strong),
            ])).Code);
    }

    [Fact]
    public void CuratedRecipesAcceptOnlyTheirDocumentedSemanticTuples()
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var caller = new CallerContext("user-a", "conversation-a", null);
        var prepared = repository.Prepare(caller,
        [
            new VisualObjectBrief(
                1,
                VisualObjectPurpose.Direction,
                VisualObjectArchetype.Arrow,
                PlacementRole: VisualObjectPlacementRole.ContentConnector,
                Recipe: VisualObjectRecipe.DirectionalCue),
            new VisualObjectBrief(
                2,
                VisualObjectPurpose.Growth,
                VisualObjectArchetype.Arrow,
                Orientation: VisualObjectOrientation.Up,
                PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                Recipe: VisualObjectRecipe.GrowthPath),
            new VisualObjectBrief(
                3,
                VisualObjectPurpose.Emphasis,
                VisualObjectArchetype.Frame,
                PlacementRole: VisualObjectPlacementRole.FocusFrame,
                Recipe: VisualObjectRecipe.FocusCorners),
            new VisualObjectBrief(
                4,
                VisualObjectPurpose.Annotation,
                VisualObjectArchetype.Callout,
                PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                Label: "Decision",
                Recipe: VisualObjectRecipe.AnnotationPin,
                AnchorCategoryOrdinal: 3,
                AnchorSeriesOrdinal: 1),
            new VisualObjectBrief(
                5,
                VisualObjectPurpose.Emphasis,
                VisualObjectArchetype.Ribbon,
                PlacementRole: VisualObjectPlacementRole.SectionDivider,
                Recipe: VisualObjectRecipe.SectionRule),
            new VisualObjectBrief(
                6,
                VisualObjectPurpose.Cycle,
                VisualObjectArchetype.Ring,
                Orientation: VisualObjectOrientation.Clockwise,
                PlacementRole: VisualObjectPlacementRole.BackgroundMotif,
                Recipe: VisualObjectRecipe.CycleCue),
        ]);

        Assert.Equal(6, prepared.Objects.Count);
        Assert.Equal(VisualObjectRecipe.DirectionalCue, prepared.Objects[0].Recipe);
        Assert.Contains("DirectionalCue", prepared.Objects[0].PreviewDescription, StringComparison.Ordinal);

        var labeledContentConnector = Assert.Throws<PptxValidationException>(() => repository.Prepare(caller,
        [
            new VisualObjectBrief(
                7,
                VisualObjectPurpose.Direction,
                VisualObjectArchetype.Arrow,
                PlacementRole: VisualObjectPlacementRole.ContentConnector,
                Label: "Duplicate relationship label",
                Recipe: VisualObjectRecipe.DirectionalCue),
        ]));
        Assert.Equal("visual_object_recipe_invalid", labeledContentConnector.Code);
    }

    [Theory]
    [InlineData(null, 1)]
    [InlineData(1, null)]
    [InlineData(0, 1)]
    [InlineData(13, 1)]
    [InlineData(1, 0)]
    [InlineData(1, 5)]
    public void AnnotationPinRequiresBoundedCategoryAndSeriesOrdinals(
        int? categoryOrdinal,
        int? seriesOrdinal)
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var exception = Assert.Throws<PptxValidationException>(() => repository.Prepare(
            new CallerContext("user-a", "conversation-a", null),
            [
                new VisualObjectBrief(
                    1,
                    VisualObjectPurpose.Annotation,
                    VisualObjectArchetype.Callout,
                    PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                    Label: "Decision",
                    Recipe: VisualObjectRecipe.AnnotationPin,
                    AnchorCategoryOrdinal: categoryOrdinal,
                    AnchorSeriesOrdinal: seriesOrdinal),
            ]));

        Assert.Equal("visual_object_recipe_invalid", exception.Code);
    }

    [Fact]
    public void OtherRecipesRejectChartAnchorOrdinals()
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var exception = Assert.Throws<PptxValidationException>(() => repository.Prepare(
            new CallerContext("user-a", "conversation-a", null),
            [
                new VisualObjectBrief(
                    1,
                    VisualObjectPurpose.Growth,
                    VisualObjectArchetype.Arrow,
                    PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                    Recipe: VisualObjectRecipe.GrowthPath,
                    AnchorCategoryOrdinal: 2,
                    AnchorSeriesOrdinal: 1),
            ]));

        Assert.Equal("visual_object_recipe_invalid", exception.Code);
    }

    [Theory]
    [InlineData(VisualObjectRecipe.DirectionalCue)]
    [InlineData(VisualObjectRecipe.GrowthPath)]
    [InlineData(VisualObjectRecipe.FocusCorners)]
    [InlineData(VisualObjectRecipe.AnnotationPin)]
    [InlineData(VisualObjectRecipe.SectionRule)]
    [InlineData(VisualObjectRecipe.CycleCue)]
    public void CuratedRecipeMismatchIsRejected(VisualObjectRecipe recipe)
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var caller = new CallerContext("user-a", "conversation-a", null);
        var exception = Assert.Throws<PptxValidationException>(() => repository.Prepare(caller,
        [
            new VisualObjectBrief(
                1,
                VisualObjectPurpose.Direction,
                VisualObjectArchetype.Arrow,
                PlacementRole: VisualObjectPlacementRole.FocusFrame,
                Recipe: recipe),
        ]));

        Assert.Equal("visual_object_recipe_invalid", exception.Code);
    }

    [Theory]
    [InlineData(VisualObjectRecipe.DirectionalCue, VisualObjectPurpose.Direction, VisualObjectArchetype.Arrow, VisualObjectPlacementRole.ContentConnector)]
    [InlineData(VisualObjectRecipe.GrowthPath, VisualObjectPurpose.Growth, VisualObjectArchetype.Arrow, VisualObjectPlacementRole.ChartAnnotation)]
    [InlineData(VisualObjectRecipe.AnnotationPin, VisualObjectPurpose.Annotation, VisualObjectArchetype.Callout, VisualObjectPlacementRole.ChartAnnotation)]
    [InlineData(VisualObjectRecipe.SectionRule, VisualObjectPurpose.Annotation, VisualObjectArchetype.Ribbon, VisualObjectPlacementRole.SectionDivider)]
    [InlineData(VisualObjectRecipe.CycleCue, VisualObjectPurpose.Cycle, VisualObjectArchetype.Ring, VisualObjectPlacementRole.BackgroundMotif)]
    public void StrongCuratedRecipeIsRejectedBeforeBrandBindingUnlessItIsFocalEmphasis(
        VisualObjectRecipe recipe,
        VisualObjectPurpose purpose,
        VisualObjectArchetype archetype,
        VisualObjectPlacementRole placementRole)
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var caller = new CallerContext("user-a", "conversation-a", null);
        var exception = Assert.Throws<PptxValidationException>(() => repository.Prepare(caller,
        [
            new VisualObjectBrief(
                1,
                purpose,
                archetype,
                Emphasis: VisualObjectEmphasis.Strong,
                Orientation: recipe == VisualObjectRecipe.CycleCue
                    ? VisualObjectOrientation.Clockwise
                    : VisualObjectOrientation.Right,
                PlacementRole: placementRole,
                Label: recipe == VisualObjectRecipe.AnnotationPin ? "Decision" : null,
                Recipe: recipe),
        ]));

        Assert.Equal("visual_object_emphasis_invalid", exception.Code);
    }

    [Fact]
    public void StrongFocusCornersWithFocalEmphasisIsAccepted()
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var prepared = repository.Prepare(
            new CallerContext("user-a", "conversation-a", null),
            [
                new VisualObjectBrief(
                    1,
                    VisualObjectPurpose.Emphasis,
                    VisualObjectArchetype.Frame,
                    Emphasis: VisualObjectEmphasis.Strong,
                    PlacementRole: VisualObjectPlacementRole.FocusFrame,
                    Recipe: VisualObjectRecipe.FocusCorners),
            ]);

        Assert.Equal(VisualObjectRecipe.FocusCorners, Assert.Single(prepared.Objects).Recipe);
    }

    [Fact]
    public void ToolResultKeepsIdsInStructuredTextAndDoesNotEmitUnsupportedImageContent()
    {
        var repository = new VisualObjectAssetRepository(TimeProvider.System);
        var prepared = repository.Prepare(
            new CallerContext("user-a", "conversation-a", null),
            [
                new VisualObjectBrief(
                    1,
                    VisualObjectPurpose.Annotation,
                    VisualObjectArchetype.Callout,
                    PlacementRole: VisualObjectPlacementRole.ChartAnnotation,
                    Label: "Decision"),
            ]);

        var result = VisualObjectPreviewResource.Create(prepared);

        Assert.False(result.IsError);
        var content = Assert.Single(result.Content);
        var text = Assert.IsType<TextContentBlock>(content).Text;
        Assert.Contains(prepared.Objects[0].AssetId, text, StringComparison.Ordinal);
        Assert.DoesNotContain("<svg", text, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class TestTimeProvider(DateTimeOffset now) : TimeProvider
    {
        private DateTimeOffset now = now;

        public override DateTimeOffset GetUtcNow() => now;

        public void Advance(TimeSpan duration) => now = now.Add(duration);
    }
}
