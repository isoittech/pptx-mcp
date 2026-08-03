using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Security;
using PptxMcp.Storage;

namespace PptxMcp.Jobs;

public sealed class JobService(
    FileJobRepository repository,
    InputFileResolver inputFileResolver,
    PptxPackageGuard packageGuard,
    JobChannel queue,
    JobCancellationRegistry cancellationRegistry,
    ArtifactTokenService tokenService,
    IOptions<PptxMcpOptions> options,
    TimeProvider timeProvider)
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);
    private readonly PptxMcpOptions options = options.Value;

    public Task<JobReceipt> SubmitAnalyzeAsync(
        CallerContext caller,
        string sourceFileId,
        CancellationToken cancellationToken) =>
        SubmitAsync<object>(caller, sourceFileId, JobKind.Analyze, payload: null, cancellationToken);

    public Task<JobReceipt> SubmitRenderAsync(
        CallerContext caller,
        string sourceFileId,
        CancellationToken cancellationToken) =>
        SubmitAsync<object>(caller, sourceFileId, JobKind.RenderPreview, payload: null, cancellationToken);

    public Task<JobReceipt> SubmitReplaceTextAsync(
        CallerContext caller,
        string sourceFileId,
        IReadOnlyList<TextReplacement> replacements,
        CancellationToken cancellationToken) =>
        SubmitAsync(caller, sourceFileId, JobKind.ReplaceText, replacements, cancellationToken);

    public Task<JobReceipt> SubmitPopulateTemplateAsync(
        CallerContext caller,
        string sourceFileId,
        IReadOnlyList<TemplateField> fields,
        CancellationToken cancellationToken) =>
        SubmitAsync(caller, sourceFileId, JobKind.PopulateTemplate, fields, cancellationToken);

    public Task<JobReceipt> SubmitCreateDeckAsync(
        CallerContext caller,
        string sourceFileId,
        IReadOnlyList<DeckSlideSpec> slides,
        CancellationToken cancellationToken) =>
        SubmitAsync(caller, sourceFileId, JobKind.CreateDeck, slides, cancellationToken);

    public async Task<JobView> GetAsync(CallerContext caller, string jobId, CancellationToken cancellationToken)
    {
        var job = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        var links = job.Artifacts.Select(artifact => CreateLink(job.Id, artifact)).ToArray();
        return new JobView(
            job.Id,
            job.Kind,
            job.State,
            job.ProgressPercent,
            job.CreatedAt,
            job.CompletedAt,
            job.Result,
            links,
            job.ErrorCode,
            job.ErrorMessage);
    }

    public async Task<bool> CancelAsync(CallerContext caller, string jobId, CancellationToken cancellationToken)
    {
        var job = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        if (job.State is JobState.Succeeded or JobState.Failed or JobState.Canceled)
        {
            return false;
        }

        await repository.UpdateAsync(
            job.Id,
            current => current with
            {
                State = JobState.Canceled,
                CompletedAt = timeProvider.GetUtcNow(),
                ErrorCode = "canceled_by_user",
                ErrorMessage = "The job was canceled by the requesting user.",
            },
            cancellationToken).ConfigureAwait(false);
        cancellationRegistry.Cancel(job.Id);
        return true;
    }

    private async Task<JobReceipt> SubmitAsync<TPayload>(
        CallerContext caller,
        string sourceFileId,
        JobKind kind,
        TPayload? payload,
        CancellationToken cancellationToken)
    {
        var input = await inputFileResolver.ResolveAsync(caller, sourceFileId, cancellationToken).ConfigureAwait(false);
        var now = timeProvider.GetUtcNow();
        var job = new JobRecord
        {
            Id = Guid.NewGuid().ToString("N"),
            Kind = kind,
            State = JobState.Queued,
            UserScope = caller.UserScope,
            ConversationScope = caller.ConversationScope,
            SourceFileId = input.FileId,
            CreatedAt = now,
            ExpiresAt = now.AddDays(options.RetentionDays),
            ProgressPercent = 0,
            Payload = payload is null ? null : JsonSerializer.SerializeToElement(payload, SerializerOptions),
        };

        try
        {
            await repository.CreateAsync(job, cancellationToken).ConfigureAwait(false);
            var sourceCopy = Path.Combine(repository.GetJobDirectory(job.Id), "source.pptx");
            await CopyFileAsync(input.Path, sourceCopy, options.MaxFileBytes, cancellationToken).ConfigureAwait(false);
            await packageGuard.ValidateAsync(sourceCopy, cancellationToken).ConfigureAwait(false);
            if (!queue.TryEnqueue(job.Id))
            {
                throw new PptxValidationException("queue_full", "The PowerPoint job queue is full. Retry later.");
            }
        }
        catch
        {
            repository.DeleteFiles(job.Id);
            throw;
        }

        return new JobReceipt(job.Id, "queued", 2);
    }

    private async Task<JobRecord> GetOwnedAsync(
        CallerContext caller,
        string jobId,
        CancellationToken cancellationToken)
    {
        JobRecord? job;
        try
        {
            job = await repository.GetAsync(jobId, cancellationToken).ConfigureAwait(false);
        }
        catch (ArgumentException)
        {
            job = null;
        }

        if (job is null
            || !string.Equals(job.UserScope, caller.UserScope, StringComparison.Ordinal)
            || !string.Equals(job.ConversationScope, caller.ConversationScope, StringComparison.Ordinal))
        {
            throw new PptxValidationException("job_not_found", "The job was not found in this conversation.");
        }

        return job;
    }

    private ArtifactLink CreateLink(string jobId, ArtifactRecord artifact)
    {
        var (token, expiresAt) = tokenService.Create(jobId, artifact.FileName);
        var escapedPath = string.Join('/', artifact.FileName.Split('/').Select(Uri.EscapeDataString));
        var baseUrl = options.PublicBaseUrl.TrimEnd('/');
        return new ArtifactLink(
            artifact.FileName,
            artifact.MediaType,
            artifact.Bytes,
            $"{baseUrl}/artifacts/{jobId}/{escapedPath}?token={Uri.EscapeDataString(token)}",
            expiresAt);
    }

    private static async Task CopyFileAsync(
        string sourcePath,
        string destinationPath,
        long maximumBytes,
        CancellationToken cancellationToken)
    {
        await using var source = File.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        await using var destination = File.Open(destinationPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        var buffer = new byte[81_920];
        long copied = 0;
        while (true)
        {
            var read = await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                break;
            }

            copied = checked(copied + read);
            if (copied > maximumBytes)
            {
                throw new PptxValidationException("file_size_out_of_range", $"PPTX files must not exceed {maximumBytes} bytes.");
            }

            await destination.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
        }
    }
}
