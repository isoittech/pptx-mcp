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
