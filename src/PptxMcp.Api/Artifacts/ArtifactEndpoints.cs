using PptxMcp.Domain;
using PptxMcp.Security;
using PptxMcp.Storage;

namespace PptxMcp.Artifacts;

public static class ArtifactEndpoints
{
    public static void MapArtifactEndpoints(this WebApplication app)
    {
        app.MapGet("/artifacts/{jobId}/{**fileName}", DownloadAsync)
            .WithName("DownloadPowerPointArtifact")
            .ExcludeFromDescription();
    }

    private static async Task<IResult> DownloadAsync(
        string jobId,
        string fileName,
        string? token,
        ArtifactTokenService tokenService,
        FileJobRepository repository,
        RetentionPolicy retentionPolicy,
        TimeProvider timeProvider,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(token) || !tokenService.Validate(jobId, fileName, token))
        {
            return Results.Unauthorized();
        }

        JobRecord? job;
        try
        {
            job = await repository.GetAsync(jobId, cancellationToken).ConfigureAwait(false);
        }
        catch (ArgumentException)
        {
            return Results.NotFound();
        }

        var artifact = job?.Artifacts.SingleOrDefault(candidate =>
            string.Equals(candidate.FileName, fileName, StringComparison.Ordinal));
        if (job is null || artifact is null || job.State != JobState.Succeeded)
        {
            return Results.NotFound();
        }

        var now = timeProvider.GetUtcNow();
        if (retentionPolicy.EffectiveExpiry(job) <= now)
        {
            return Results.StatusCode(StatusCodes.Status410Gone);
        }

        var jobDirectory = repository.GetJobDirectory(job.Id);
        var path = Path.GetFullPath(Path.Combine(jobDirectory, fileName));
        var directoryPrefix = Path.GetFullPath(jobDirectory) + Path.DirectorySeparatorChar;
        if (!path.StartsWith(directoryPrefix, StringComparison.Ordinal) || !File.Exists(path))
        {
            return Results.NotFound();
        }

        FileStream artifactStream;
        try
        {
            artifactStream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        }
        catch (IOException)
        {
            return Results.NotFound();
        }

        try
        {
            if (artifact.StartsDownloadRetention && job.FirstDownloadedAt is null)
            {
                await repository.UpdateAsync(
                    job.Id,
                    current => current.FirstDownloadedAt is null
                        ? current with { FirstDownloadedAt = now }
                        : current,
                    cancellationToken).ConfigureAwait(false);
            }
        }
        catch
        {
            await artifactStream.DisposeAsync().ConfigureAwait(false);
            throw;
        }

        return Results.File(
            artifactStream,
            artifact.MediaType,
            fileDownloadName: artifact.StartsDownloadRetention ? "presentation.pptx" : null,
            enableRangeProcessing: true);
    }
}
