using PptxMcp.Artifacts;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Jobs;

public sealed class RetentionWorker(
    FileJobRepository repository,
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
    }
}
