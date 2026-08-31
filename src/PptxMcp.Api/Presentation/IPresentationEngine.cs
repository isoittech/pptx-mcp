using PptxMcp.Domain;

namespace PptxMcp.Presentation;

public interface IPresentationEngine
{
    Task<PresentationSummary> AnalyzeAsync(string sourcePath, CancellationToken cancellationToken);

    Task<EditResult> ReplaceTextAsync(
        string sourcePath,
        string destinationPath,
        IReadOnlyList<TextReplacement> replacements,
        CancellationToken cancellationToken);

    Task<EditResult> PopulateTemplateAsync(
        string sourcePath,
        string destinationPath,
        IReadOnlyList<TemplateField> fields,
        CancellationToken cancellationToken);

    Task<DeckCreationResult> CreateDeckAsync(
        string sourcePath,
        string destinationPath,
        IReadOnlyList<DeckSlideSpec> slides,
        CancellationToken cancellationToken);

    Task<BrandedVisualCompositionResult> CreateBrandedVisualDeckAsync(
        string templatePath,
        string visualDeckPath,
        string destinationPath,
        string templateLayoutId,
        CancellationToken cancellationToken,
        TemplateLayoutRolePolicy? layoutRolePolicy = null);
}
