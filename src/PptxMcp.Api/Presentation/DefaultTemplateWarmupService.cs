using PptxMcp.Storage;

namespace PptxMcp.Presentation;

public sealed class DefaultTemplateWarmupService(
    TemplateRegistry templates,
    PresentationAnalysisCache analysisCache,
    ILogger<DefaultTemplateWarmupService> logger) : IHostedService
{
    private static readonly Action<ILogger, string, int, Exception?> LogTemplateReady =
        LoggerMessage.Define<string, int>(
            LogLevel.Information,
            new EventId(1101, nameof(LogTemplateReady)),
            "Default PowerPoint template {TemplateId} is ready with {SlideCount} source slides.");

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        var template = await templates.TryResolveDefaultAsync(cancellationToken).ConfigureAwait(false);
        if (template is null)
        {
            return;
        }

        var summary = await analysisCache.GetAsync(template.Path, cancellationToken).ConfigureAwait(false);
        LogTemplateReady(logger, template.FileId, summary.SlideCount, null);
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
}
