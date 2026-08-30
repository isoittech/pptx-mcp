using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Presentation;
using PptxMcp.Storage;

namespace PptxMcp.Jobs;

public sealed class JobWorker(
    JobChannel queue,
    FileJobRepository repository,
    IPresentationEngine presentationEngine,
    PresentationAnalysisCache analysisCache,
    IVisualPresentationEngine visualPresentationEngine,
    LibreOfficeRenderer renderer,
    PptxPackageGuard packageGuard,
    JobCancellationRegistry cancellationRegistry,
    IOptions<PptxMcpOptions> options,
    TimeProvider timeProvider,
    ILogger<JobWorker> logger) : BackgroundService
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);
    private static readonly Action<ILogger, string, Exception?> LogUnexpectedFailure = LoggerMessage.Define<string>(
        LogLevel.Error,
        new EventId(1001, nameof(LogUnexpectedFailure)),
        "Unexpected failure while processing job {JobId}.");
    private static readonly Action<ILogger, string, Exception?> LogJobFailure = LoggerMessage.Define<string>(
        LogLevel.Error,
        new EventId(1002, nameof(LogJobFailure)),
        "PowerPoint job {JobId} failed.");
    private readonly PptxMcpOptions options = options.Value;

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        await RestoreQueuedJobsAsync(stoppingToken).ConfigureAwait(false);
        var workers = Enumerable.Range(0, options.MaxConcurrentJobs)
            .Select(_ => RunWorkerAsync(stoppingToken))
            .ToArray();
        await Task.WhenAll(workers).ConfigureAwait(false);
    }

    private async Task RestoreQueuedJobsAsync(CancellationToken cancellationToken)
    {
        await foreach (var job in repository.ListAsync(cancellationToken))
        {
            if (job.State is not (JobState.Queued or JobState.Running))
            {
                continue;
            }

            await repository.UpdateAsync(
                job.Id,
                current => current with { State = JobState.Queued, StartedAt = null, ProgressPercent = 0 },
                cancellationToken).ConfigureAwait(false);
            if (!queue.TryEnqueue(job.Id))
            {
                await FailAsync(job.Id, "queue_full_after_restart", "The restored job queue is full.", cancellationToken)
                    .ConfigureAwait(false);
            }
        }
    }

    private async Task RunWorkerAsync(CancellationToken stoppingToken)
    {
        await foreach (var jobId in queue.ReadAllAsync(stoppingToken))
        {
            try
            {
                await ProcessAsync(jobId, stoppingToken).ConfigureAwait(false);
            }
            catch (Exception exception) when (exception is not OperationCanceledException || !stoppingToken.IsCancellationRequested)
            {
                LogUnexpectedFailure(logger, jobId, exception);
            }
        }
    }

    private async Task ProcessAsync(string jobId, CancellationToken stoppingToken)
    {
        var job = await repository.GetAsync(jobId, stoppingToken).ConfigureAwait(false);
        if (job is null || job.State != JobState.Queued)
        {
            return;
        }

        await repository.UpdateAsync(
            jobId,
            current => current with
            {
                State = JobState.Running,
                StartedAt = timeProvider.GetUtcNow(),
                ProgressPercent = 5,
            },
            stoppingToken).ConfigureAwait(false);

        using var lease = cancellationRegistry.Register(
            jobId,
            TimeSpan.FromMinutes(options.JobTimeoutMinutes),
            stoppingToken);

        try
        {
            var directory = repository.GetJobDirectory(jobId);
            var sourcePath = Path.Combine(directory, "source.pptx");
            var (result, artifacts) = await ExecuteJobAsync(job, sourcePath, directory, lease.Token).ConfigureAwait(false);
            await repository.UpdateAsync(
                jobId,
                current => current.State == JobState.Canceled
                    ? current
                    : current with
                    {
                        State = JobState.Succeeded,
                        ProgressPercent = 100,
                        CompletedAt = timeProvider.GetUtcNow(),
                        Result = result,
                        Artifacts = artifacts,
                    },
                stoppingToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (!stoppingToken.IsCancellationRequested)
        {
            var current = await repository.GetAsync(jobId, CancellationToken.None).ConfigureAwait(false);
            if (current?.State != JobState.Canceled)
            {
                await FailAsync(jobId, "job_timeout", "The job exceeded its execution time limit.", CancellationToken.None)
                    .ConfigureAwait(false);
            }
        }
        catch (PptxValidationException exception)
        {
            await FailAsync(jobId, exception.Code, exception.Message, stoppingToken).ConfigureAwait(false);
        }
        catch (Exception exception)
        {
            LogJobFailure(logger, jobId, exception);
            await FailAsync(jobId, "internal_error", "The PowerPoint job failed unexpectedly.", stoppingToken)
                .ConfigureAwait(false);
        }
    }

    private async Task<(JsonElement? Result, IReadOnlyList<ArtifactRecord> Artifacts)> ExecuteJobAsync(
        JobRecord job,
        string sourcePath,
        string directory,
        CancellationToken cancellationToken)
    {
        switch (job.Kind)
        {
            case JobKind.Analyze:
                {
                    var summary = await analysisCache.GetAsync(sourcePath, cancellationToken).ConfigureAwait(false);
                    var payload = job.Payload?.Deserialize<AnalyzeJobPayload>(SerializerOptions);
                    var result = PrepareAnalysisResult(summary, payload?.IncludeLayouts ?? false);
                    return (JsonSerializer.SerializeToElement(result, SerializerOptions), []);
                }

            case JobKind.RenderPreview:
                {
                    var images = await RenderAsync(sourcePath, directory, cancellationToken).ConfigureAwait(false);
                    return (JsonSerializer.SerializeToElement(new { preview_count = images.Count }), CreateImageArtifacts(images, directory));
                }

            case JobKind.ReplaceText:
                {
                    var payload = DeserializeReplaceTextPayload(job.Payload);
                    var outputPath = Path.Combine(directory, "presentation.pptx");
                    var edit = await presentationEngine.ReplaceTextAsync(
                            sourcePath,
                            outputPath,
                            payload.Replacements,
                            cancellationToken)
                        .ConfigureAwait(false);
                    await packageGuard.ValidateAsync(outputPath, cancellationToken).ConfigureAwait(false);
                    var images = payload.IsFinalBatch
                        ? await RenderAsync(outputPath, directory, cancellationToken).ConfigureAwait(false)
                        : [];
                    return (
                        JsonSerializer.SerializeToElement(edit, SerializerOptions),
                        CreateOutputArtifacts(outputPath, images, directory));
                }

            case JobKind.PopulateTemplate:
                {
                    var fields = job.Payload?.Deserialize<List<TemplateField>>(SerializerOptions)
                        ?? throw new PptxValidationException("invalid_job_payload", "Template fields are missing.");
                    var outputPath = Path.Combine(directory, "presentation.pptx");
                    var edit = await presentationEngine.PopulateTemplateAsync(sourcePath, outputPath, fields, cancellationToken)
                        .ConfigureAwait(false);
                    await packageGuard.ValidateAsync(outputPath, cancellationToken).ConfigureAwait(false);
                    var images = await RenderAsync(outputPath, directory, cancellationToken).ConfigureAwait(false);
                    return (
                        JsonSerializer.SerializeToElement(edit, SerializerOptions),
                        CreateOutputArtifacts(outputPath, images, directory));
                }

            case JobKind.CreateDeck:
                {
                    var slides = job.Payload?.Deserialize<List<DeckSlideSpec>>(SerializerOptions)
                        ?? throw new PptxValidationException("invalid_job_payload", "Deck slide specifications are missing.");
                    var outputPath = Path.Combine(directory, "presentation.pptx");
                    var creation = await presentationEngine.CreateDeckAsync(sourcePath, outputPath, slides, cancellationToken)
                        .ConfigureAwait(false);
                    await packageGuard.ValidateAsync(outputPath, cancellationToken).ConfigureAwait(false);
                    var images = await RenderAsync(outputPath, directory, cancellationToken).ConfigureAwait(false);
                    return (
                        JsonSerializer.SerializeToElement(creation, SerializerOptions),
                        CreateOutputArtifacts(outputPath, images, directory));
                }

            case JobKind.CreateVisualDeck:
                {
                    var deck = job.Payload?.Deserialize<VisualDeckSpec>(SerializerOptions)
                        ?? throw new PptxValidationException("invalid_job_payload", "Visual deck specification is missing.");
                    VisualDeckValidator.Validate(deck, options.MaxSlides);
                    var outputPath = Path.Combine(directory, "presentation.pptx");
                    var creation = await visualPresentationEngine.CreateAsync(outputPath, deck, false, cancellationToken)
                        .ConfigureAwait(false);
                    await packageGuard.ValidateAsync(outputPath, cancellationToken).ConfigureAwait(false);
                    var images = await RenderAsync(outputPath, directory, cancellationToken).ConfigureAwait(false);
                    return (
                        JsonSerializer.SerializeToElement(creation, SerializerOptions),
                        CreateOutputArtifacts(outputPath, images, directory));
                }

            case JobKind.CreateBrandedVisualDeck:
                {
                    var branded = job.Payload?.Deserialize<BrandedVisualDeckSpec>(SerializerOptions)
                        ?? throw new PptxValidationException("invalid_job_payload", "Branded visual deck specifications are missing.");
                    VisualDeckValidator.Validate(branded.Deck, options.MaxSlides);
                    var templateSummary = await analysisCache.GetAsync(sourcePath, cancellationToken).ConfigureAwait(false);
                    var themedDeck = VisualDeckBranding.ApplyTemplateTheme(branded.Deck, templateSummary.Theme);
                    var visualPath = Path.Combine(directory, "visual-source.pptx");
                    var visualCreation = await visualPresentationEngine.CreateAsync(
                            visualPath,
                            themedDeck,
                            true,
                            cancellationToken)
                        .ConfigureAwait(false);
                    await packageGuard.ValidateAsync(visualPath, cancellationToken).ConfigureAwait(false);

                    var outputPath = Path.Combine(directory, "presentation.pptx");
                    var defaultTemplateRolePolicy = string.Equals(
                            job.SourceFileId,
                            options.DefaultTemplateId,
                            StringComparison.Ordinal)
                        && options.DefaultTemplateCoverSampleSlideNumber > 0
                        && options.DefaultTemplateBodySampleSlideNumber > 0
                        ? new TemplateLayoutRolePolicy(
                            options.DefaultTemplateCoverSampleSlideNumber,
                            options.DefaultTemplateBodySampleSlideNumber)
                        : null;
                    var composition = await presentationEngine.CreateBrandedVisualDeckAsync(
                            sourcePath,
                            visualPath,
                            outputPath,
                            branded.TemplateLayoutId,
                            cancellationToken,
                            defaultTemplateRolePolicy)
                        .ConfigureAwait(false);
                    await packageGuard.ValidateAsync(outputPath, cancellationToken).ConfigureAwait(false);
                    var images = await RenderAsync(outputPath, directory, cancellationToken).ConfigureAwait(false);
                    var result = new BrandedVisualDeckCreationResult(
                        composition.SlideCount,
                        visualCreation.LayoutKinds,
                        visualCreation.Renderer,
                        composition.TemplateLayoutId,
                        composition.TemplateLayoutName,
                        templateSummary.Theme is not null,
                        visualCreation.SpeakerNotesCount,
                        visualCreation.DesignWarnings);
                    return (
                        JsonSerializer.SerializeToElement(result, SerializerOptions),
                        CreateOutputArtifacts(outputPath, images, directory));
                }

            default:
                throw new PptxValidationException("unsupported_job", $"Job type '{job.Kind}' is not supported.");
        }
    }

    internal static object PrepareAnalysisResult(
        PresentationSummary summary,
        bool includeLayouts) => includeLayouts
        ? summary
        : new TextEditAnalysisResult(
            summary.AnalysisTruncated,
            summary.HasCharts,
            summary.Slides.Select(static slide => new object[]
            {
                slide.SlideNumber,
                slide.Shapes
                    .SelectMany(static shape => shape.ExactTexts
                        ?? (string.IsNullOrWhiteSpace(shape.Text) ? [] : [shape.Text!]))
                    .Where(static text => !string.IsNullOrWhiteSpace(text))
                    .ToArray(),
            }).ToArray(),
            summary.ValidationErrors.Count > 0 ? summary.ValidationErrors : null);

    internal static ReplaceTextJobPayload DeserializeReplaceTextPayload(JsonElement? serialized)
    {
        if (serialized is null)
        {
            throw new PptxValidationException("invalid_job_payload", "Replacement instructions are missing.");
        }

        if (serialized.Value.ValueKind == JsonValueKind.Array)
        {
            var legacyReplacements = serialized.Value.Deserialize<List<TextReplacement>>(SerializerOptions)
                ?? throw new PptxValidationException("invalid_job_payload", "Replacement instructions are missing.");
            return new ReplaceTextJobPayload(legacyReplacements, IsFinalBatch: true);
        }

        return serialized.Value.Deserialize<ReplaceTextJobPayload>(SerializerOptions)
            ?? throw new PptxValidationException("invalid_job_payload", "Replacement instructions are missing.");
    }

    private async Task<IReadOnlyList<string>> RenderAsync(
        string presentationPath,
        string jobDirectory,
        CancellationToken cancellationToken)
    {
        var previewDirectory = Path.Combine(jobDirectory, "preview");
        return await renderer.RenderAsync(presentationPath, previewDirectory, cancellationToken).ConfigureAwait(false);
    }

    private static List<ArtifactRecord> CreateOutputArtifacts(
        string outputPath,
        IReadOnlyList<string> images,
        string jobDirectory)
    {
        var artifacts = new List<ArtifactRecord>
        {
            new("presentation.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", new FileInfo(outputPath).Length, true),
        };
        artifacts.AddRange(CreateImageArtifacts(images, jobDirectory));
        return artifacts;
    }

    private static ArtifactRecord[] CreateImageArtifacts(
        IReadOnlyList<string> images,
        string jobDirectory) =>
        images.Select(path => new ArtifactRecord(
                Path.GetRelativePath(jobDirectory, path).Replace('\\', '/'),
                "image/png",
                new FileInfo(path).Length,
                false))
            .ToArray();

    private Task<JobRecord> FailAsync(string jobId, string code, string message, CancellationToken cancellationToken) =>
        repository.UpdateAsync(
            jobId,
            current => current.State == JobState.Canceled
                ? current
                : current with
                {
                    State = JobState.Failed,
                    CompletedAt = timeProvider.GetUtcNow(),
                    ErrorCode = code,
                    ErrorMessage = message,
                },
            cancellationToken);
}
