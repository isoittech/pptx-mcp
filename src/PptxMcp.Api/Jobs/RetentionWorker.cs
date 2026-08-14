using PptxMcp.Artifacts;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Jobs;

public sealed class RetentionWorker(
    FileJobRepository repository,
    ImageAssetRepository imageAssets,
    RetentionPolicy retentionPolicy,
    TimeProvider timeProvider,
    ILogger<RetentionWorker> logger) : BackgroundService
{
    private static readonly Action<ILogger, string, Exception?> LogDeleted = LoggerMessage.Define<string>(
        LogLevel.Information,
        new EventId(2001, nameof(LogDeleted)),
        "Deleted expired PowerPoint job {JobId}.");
    private static readonly Action<ILogger, string, Exception?> LogDeleteFailure = LoggerMessage.Define<string>(
        LogLevel.Warning,
        new EventId(2002, nameof(LogDeleteFailure)),
        "Could not delete expired PowerPoint job {JobId}.");
    private static readonly Action<ILogger, string, Exception?> LogDeletedImageAsset = LoggerMessage.Define<string>(
        LogLevel.Information,
        new EventId(2003, nameof(LogDeletedImageAsset)),
        "Deleted expired PowerPoint image asset {AssetId}.");
    private static readonly Action<ILogger, string, Exception?> LogImageAssetDeleteFailure = LoggerMessage.Define<string>(
        LogLevel.Warning,
        new EventId(2004, nameof(LogImageAssetDeleteFailure)),
        "Could not delete expired PowerPoint image asset {AssetId}.");
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        await SweepAsync(stoppingToken).ConfigureAwait(false);
        using var timer = new PeriodicTimer(TimeSpan.FromHours(1), timeProvider);
        while (await timer.WaitForNextTickAsync(stoppingToken).ConfigureAwait(false))
        {
            await SweepAsync(stoppingToken).ConfigureAwait(false);
        }
    }

    private async Task SweepAsync(CancellationToken cancellationToken)
    {
        var now = timeProvider.GetUtcNow();
        await foreach (var job in repository.ListAsync(cancellationToken))
        {
            if (retentionPolicy.EffectiveExpiry(job) > now || job.State == JobState.Running)
            {
                continue;
            }

            try
            {
                repository.DeleteFiles(job.Id);
                LogDeleted(logger, job.Id, null);
            }
            catch (IOException exception)
            {
                LogDeleteFailure(logger, job.Id, exception);
            }
        }

        foreach (var asset in imageAssets.List())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (asset.ExpiresAt > now)
            {
                continue;
            }

            try
            {
                imageAssets.Delete(asset.AssetId);
                LogDeletedImageAsset(logger, asset.AssetId, null);
            }
            catch (IOException exception)
            {
                LogImageAssetDeleteFailure(logger, asset.AssetId, exception);
            }
        }
    }
}
