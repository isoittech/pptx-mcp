namespace PptxMcp.Design;

public sealed class BrandProfileCatalogWarmupService(
    BrandProfileCatalog catalog,
    ILogger<BrandProfileCatalogWarmupService> logger) : IHostedService
{
    private static readonly Action<ILogger, Exception?> LogCatalogReady =
        LoggerMessage.Define(
            LogLevel.Information,
            new EventId(1201, nameof(LogCatalogReady)),
            "External Brand Profile catalog is ready.");

    public Task StartAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        catalog.EnsureReady();
        LogCatalogReady(logger, null);
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
}
