using PptxMcp.Domain;
using PptxMcp.Presentation;

namespace PptxMcp.Tests;

public sealed class PresentationAnalysisCacheTests
{
    [Fact]
    public async Task ReusesAnalysisForIdenticalTemplateContent()
    {
        var firstPath = TestPresentationFactory.Create("cache");
        var secondPath = Path.Combine(Path.GetTempPath(), $"pptx-cache-copy-{Guid.NewGuid():N}.pptx");
        File.Copy(firstPath, secondPath);
        var engine = new CountingPresentationEngine();
        var cache = new PresentationAnalysisCache(engine);

        try
        {
            var first = await cache.GetAsync(firstPath, CancellationToken.None);
            var second = await cache.GetAsync(secondPath, CancellationToken.None);

            Assert.Same(first, second);
            Assert.Equal(1, engine.AnalysisCount);
        }
        finally
        {
            File.Delete(firstPath);
            File.Delete(secondPath);
        }
    }

    private sealed class CountingPresentationEngine : IPresentationEngine
    {
        private int analysisCount;

        public int AnalysisCount => analysisCount;

        public Task<PresentationSummary> AnalyzeAsync(string sourcePath, CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref analysisCount);
            return Task.FromResult(new PresentationSummary(1, false, false, false, false, [], [], null, []));
        }

        public Task<EditResult> ReplaceTextAsync(
            string sourcePath,
            string destinationPath,
            IReadOnlyList<TextReplacement> replacements,
            CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<EditResult> PopulateTemplateAsync(
            string sourcePath,
            string destinationPath,
            IReadOnlyList<TemplateField> fields,
            CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<DeckCreationResult> CreateDeckAsync(
            string sourcePath,
            string destinationPath,
            IReadOnlyList<DeckSlideSpec> slides,
            CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<BrandedVisualCompositionResult> CreateBrandedVisualDeckAsync(
            string templatePath,
            string visualDeckPath,
            string destinationPath,
            string templateLayoutId,
            CancellationToken cancellationToken,
            TemplateLayoutRolePolicy? layoutRolePolicy = null) => throw new NotSupportedException();
    }
}
