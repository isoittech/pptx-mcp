using PptxMcp.Domain;

namespace PptxMcp.Presentation;

public interface IVisualPresentationEngine
{
    Task<VisualDeckCreationResult> CreateAsync(
        string destinationPath,
        VisualDeckSpec deck,
        bool useTemplateChrome,
        bool useDefaultTemplateCoverOverlay,
        bool useDefaultTemplateBodyStyle,
        CancellationToken cancellationToken);
}
