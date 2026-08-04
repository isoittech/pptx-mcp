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

    public async Task<JobReceipt> SubmitRefineDeckAsync(
        CallerContext caller,
        string jobId,
        IReadOnlyList<DeckSlideRevision> revisions,
        CancellationToken cancellationToken)
    {
        var sourceJob = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        if (sourceJob.Kind != JobKind.CreateDeck || sourceJob.State != JobState.Succeeded)
        {
            throw new PptxValidationException(
                "deck_job_not_refinable",
                "Only a successful pptx_create_deck job can be refined.");
        }

        var originalSlides = sourceJob.Payload?.Deserialize<List<DeckSlideSpec>>(SerializerOptions)
            ?? throw new PptxValidationException("invalid_job_payload", "The source deck specification is missing.");
        var refinedSlides = ApplyDeckRevisions(originalSlides, revisions, options.MaxSlides);
        var sourcePath = Path.Combine(repository.GetJobDirectory(sourceJob.Id), "source.pptx");
        if (!File.Exists(sourcePath))
        {
            throw new PptxValidationException("source_expired", "The source template is no longer available.");
        }

        return await SubmitFromPathAsync(
            caller,
            sourceJob.SourceFileId,
            sourcePath,
            JobKind.CreateDeck,
            refinedSlides,
            cancellationToken).ConfigureAwait(false);
    }

    public Task<JobReceipt> SubmitVisualDeckAsync(
        CallerContext caller,
        VisualDeckSpec deck,
        CancellationToken cancellationToken)
    {
        VisualDeckValidator.Validate(deck, options.MaxSlides);
        return SubmitGeneratedAsync(caller, JobKind.CreateVisualDeck, deck, cancellationToken);
    }

    public async Task<JobReceipt> SubmitRefineVisualDeckAsync(
        CallerContext caller,
        string jobId,
        IReadOnlyList<VisualSlideRevision> revisions,
        CancellationToken cancellationToken)
    {
        var sourceJob = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        if (sourceJob.Kind != JobKind.CreateVisualDeck || sourceJob.State != JobState.Succeeded)
        {
            throw new PptxValidationException(
                "visual_deck_job_not_refinable",
                "Only a successful pptx_create_visual_deck job can be refined.");
        }

        var originalDeck = sourceJob.Payload?.Deserialize<VisualDeckSpec>(SerializerOptions)
            ?? throw new PptxValidationException("invalid_job_payload", "The source visual deck specification is missing.");
        var refinedDeck = ApplyVisualDeckRevisions(originalDeck, revisions, options.MaxSlides);
        VisualDeckValidator.Validate(refinedDeck, options.MaxSlides);
        return await SubmitGeneratedAsync(
            caller,
            JobKind.CreateVisualDeck,
            refinedDeck,
            cancellationToken).ConfigureAwait(false);
    }

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

    public async Task<IReadOnlyList<PreviewImageData>> GetPreviewImagesAsync(
        CallerContext caller,
        string jobId,
        IReadOnlyList<int> slideNumbers,
        CancellationToken cancellationToken)
    {
        if (slideNumbers is null || slideNumbers.Count is < 1 or > 4 || slideNumbers.Distinct().Count() != slideNumbers.Count)
        {
            throw new PptxValidationException(
                "preview_selection_invalid",
                "Select between 1 and 4 distinct slide numbers per visual review call.");
        }

        var job = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        if (job.State != JobState.Succeeded)
        {
            throw new PptxValidationException(
                "job_not_ready",
                "The job must succeed before preview images can be reviewed.");
        }

        var jobDirectory = repository.GetJobDirectory(job.Id);
        var images = new List<PreviewImageData>(slideNumbers.Count);
        foreach (var slideNumber in slideNumbers)
        {
            if (slideNumber is < 1 or > 50)
            {
                throw new PptxValidationException("preview_slide_invalid", "Slide numbers must be between 1 and 50.");
            }

            var artifact = job.Artifacts.SingleOrDefault(candidate =>
                string.Equals(candidate.MediaType, "image/png", StringComparison.Ordinal)
                && TryGetPreviewSlideNumber(candidate.FileName) == slideNumber);
            if (artifact is null)
            {
                throw new PptxValidationException(
                    "preview_slide_not_found",
                    $"Preview image for slide {slideNumber} was not found in this job.");
            }

            if (artifact.Bytes > 8L * 1024 * 1024)
            {
                throw new PptxValidationException(
                    "preview_image_too_large",
                    $"Preview image for slide {slideNumber} exceeds the visual review limit.");
            }

            var path = Path.GetFullPath(Path.Combine(jobDirectory, artifact.FileName));
            var root = Path.GetFullPath(jobDirectory) + Path.DirectorySeparatorChar;
            if (!path.StartsWith(root, StringComparison.Ordinal) || !File.Exists(path))
            {
                throw new PptxValidationException("preview_slide_not_found", "The preview image is no longer available.");
            }

            images.Add(new PreviewImageData(
                slideNumber,
                artifact.MediaType,
                await File.ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false)));
        }

        return images;
    }

    internal static int? TryGetPreviewSlideNumber(string fileName)
    {
        if (!fileName.StartsWith("preview/slide-", StringComparison.Ordinal)
            || !fileName.EndsWith(".png", StringComparison.Ordinal))
        {
            return null;
        }

        var numberStart = "preview/slide-".Length;
        var number = fileName[numberStart..^".png".Length];
        return int.TryParse(number, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    internal static IReadOnlyList<DeckSlideSpec> ApplyDeckRevisions(
        IReadOnlyList<DeckSlideSpec> originalSlides,
        IReadOnlyList<DeckSlideRevision> revisions,
        int maximumSlides)
    {
        if (originalSlides is null
            || originalSlides.Count is < 1
            || originalSlides.Count > maximumSlides
            || revisions is null
            || revisions.Count is < 1
            || revisions.Count > originalSlides.Count
            || revisions.Any(static revision => revision is null || revision.Fields is null)
            || revisions.Select(static revision => revision.SlideNumber).Distinct().Count() != revisions.Count
            || revisions.Any(revision => revision.SlideNumber < 1 || revision.SlideNumber > originalSlides.Count))
        {
            throw new PptxValidationException(
                "deck_revision_invalid",
                "Provide 1 or more distinct slide revisions within the source deck slide range.");
        }

        var revisionsBySlide = revisions.ToDictionary(static revision => revision.SlideNumber);
        return originalSlides
            .Select((slide, index) => revisionsBySlide.TryGetValue(index + 1, out var revision)
                ? new DeckSlideSpec(slide.LayoutId, revision.Fields)
                : slide)
            .ToArray();
    }

    internal static VisualDeckSpec ApplyVisualDeckRevisions(
        VisualDeckSpec originalDeck,
        IReadOnlyList<VisualSlideRevision> revisions,
        int maximumSlides)
    {
        if (originalDeck is null
            || originalDeck.Slides is null
            || originalDeck.Slides.Count is < 1
            || originalDeck.Slides.Count > maximumSlides
            || revisions is null
            || revisions.Count is < 1
            || revisions.Count > originalDeck.Slides.Count
            || revisions.Any(static revision => revision is null || revision.Slide is null)
            || revisions.Select(static revision => revision.SlideNumber).Distinct().Count() != revisions.Count
            || revisions.Any(revision => revision.SlideNumber < 1 || revision.SlideNumber > originalDeck.Slides.Count))
        {
            throw new PptxValidationException(
                "visual_deck_revision_invalid",
                "Provide 1 or more distinct visual slide revisions within the source deck slide range.");
        }

        var revisionsBySlide = revisions.ToDictionary(static revision => revision.SlideNumber);
        var slides = originalDeck.Slides
            .Select((slide, index) => revisionsBySlide.TryGetValue(index + 1, out var revision)
                ? revision.Slide
                : slide)
            .ToArray();
        return originalDeck with { Slides = slides };
    }

    private async Task<JobReceipt> SubmitAsync<TPayload>(
        CallerContext caller,
        string sourceFileId,
        JobKind kind,
        TPayload? payload,
        CancellationToken cancellationToken)
    {
        var input = await inputFileResolver.ResolveAsync(caller, sourceFileId, cancellationToken).ConfigureAwait(false);
        return await SubmitFromPathAsync(
            caller,
            input.FileId,
            input.Path,
            kind,
            payload,
            cancellationToken).ConfigureAwait(false);
    }

    private async Task<JobReceipt> SubmitFromPathAsync<TPayload>(
        CallerContext caller,
        string sourceFileId,
        string sourcePath,
        JobKind kind,
        TPayload? payload,
        CancellationToken cancellationToken)
    {
        var now = timeProvider.GetUtcNow();
        var job = new JobRecord
        {
            Id = Guid.NewGuid().ToString("N"),
            Kind = kind,
            State = JobState.Queued,
            UserScope = caller.UserScope,
            ConversationScope = caller.ConversationScope,
            SourceFileId = sourceFileId,
            CreatedAt = now,
            ExpiresAt = now.AddDays(options.RetentionDays),
            ProgressPercent = 0,
            Payload = payload is null ? null : JsonSerializer.SerializeToElement(payload, SerializerOptions),
        };

        try
        {
            await repository.CreateAsync(job, cancellationToken).ConfigureAwait(false);
            var sourceCopy = Path.Combine(repository.GetJobDirectory(job.Id), "source.pptx");
            await CopyFileAsync(sourcePath, sourceCopy, options.MaxFileBytes, cancellationToken).ConfigureAwait(false);
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

    private async Task<JobReceipt> SubmitGeneratedAsync<TPayload>(
        CallerContext caller,
        JobKind kind,
        TPayload payload,
        CancellationToken cancellationToken)
    {
        var now = timeProvider.GetUtcNow();
        var job = new JobRecord
        {
            Id = Guid.NewGuid().ToString("N"),
            Kind = kind,
            State = JobState.Queued,
            UserScope = caller.UserScope,
            ConversationScope = caller.ConversationScope,
            SourceFileId = "generated",
            CreatedAt = now,
            ExpiresAt = now.AddDays(options.RetentionDays),
            ProgressPercent = 0,
            Payload = JsonSerializer.SerializeToElement(payload, SerializerOptions),
        };

        try
        {
            await repository.CreateAsync(job, cancellationToken).ConfigureAwait(false);
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
