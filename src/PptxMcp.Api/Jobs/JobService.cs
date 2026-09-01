using System.Collections.Concurrent;
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
    TemplateRegistry templates,
    PptxPackageGuard packageGuard,
    JobChannel queue,
    JobCancellationRegistry cancellationRegistry,
    ArtifactTokenService tokenService,
    IOptions<PptxMcpOptions> options,
    TimeProvider timeProvider,
    ImageAssetRepository? imageAssets = null)
{
    public const int MaximumVisualRevisionRounds = 3;
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);
    private readonly ConcurrentDictionary<string, AnalyzeSubmission> analyzeSubmissions = new();
    private readonly ConcurrentDictionary<string, TextEditWorkflow> textEditWorkflows = new();
    private readonly SemaphoreSlim[] visualMutationLocks = Enumerable.Range(0, 64)
        .Select(static _ => new SemaphoreSlim(1, 1))
        .ToArray();
    private readonly PptxMcpOptions options = options.Value;
    private readonly ImageAssetRepository? imageAssets = imageAssets;

    public Task<JobReceipt> SubmitAnalyzeAsync(
        CallerContext caller,
        string sourceFileId,
        CancellationToken cancellationToken,
        bool includeLayouts = false)
    {
        var key = CreateAnalyzeSubmissionKey(caller, sourceFileId, includeLayouts);
        if (key is null)
        {
            return SubmitAnalyzeCoreAsync(caller, sourceFileId, includeLayouts, cancellationToken);
        }

        var now = timeProvider.GetUtcNow();
        RemoveExpiredAnalyzeSubmissions(now);
        var submission = analyzeSubmissions.GetOrAdd(
            key,
            _ => new AnalyzeSubmission(
                new Lazy<Task<JobReceipt>>(
                    () => SubmitAnalyzeCoreAsync(caller, sourceFileId, includeLayouts, cancellationToken),
                    LazyThreadSafetyMode.ExecutionAndPublication),
                now.AddHours(24)));

        return AwaitAnalyzeSubmissionAsync(key, submission);
    }

    private Task<JobReceipt> SubmitAnalyzeCoreAsync(
        CallerContext caller,
        string sourceFileId,
        bool includeLayouts,
        CancellationToken cancellationToken) => string.Equals(sourceFileId, "default", StringComparison.OrdinalIgnoreCase)
        ? SubmitTemplateAsync(caller, sourceFileId, JobKind.Analyze, new AnalyzeJobPayload(includeLayouts), cancellationToken)
        : SubmitAsync(caller, sourceFileId, JobKind.Analyze, new AnalyzeJobPayload(includeLayouts), cancellationToken);

    private async Task<JobReceipt> AwaitAnalyzeSubmissionAsync(string key, AnalyzeSubmission submission)
    {
        try
        {
            return await submission.Receipt.Value.ConfigureAwait(false);
        }
        catch
        {
            analyzeSubmissions.TryRemove(new KeyValuePair<string, AnalyzeSubmission>(key, submission));
            throw;
        }
    }

    private void RemoveExpiredAnalyzeSubmissions(DateTimeOffset now)
    {
        foreach (var entry in analyzeSubmissions)
        {
            if (entry.Value.ExpiresAt <= now)
            {
                analyzeSubmissions.TryRemove(entry);
            }
        }
    }

    public Task<JobReceipt> SubmitRenderAsync(
        CallerContext caller,
        string sourceFileId,
        CancellationToken cancellationToken)
    {
        RemoveExpiredAnalyzeSubmissions(timeProvider.GetUtcNow());
        var analysisKey = CreateAnalyzeSubmissionKey(caller, sourceFileId, includeLayouts: false);
        if (analysisKey is not null
            && analyzeSubmissions.ContainsKey(analysisKey))
        {
            throw new PptxValidationException(
                "source_preview_before_text_edit_forbidden",
                "The source presentation was already analyzed for text editing in this user message. Do not render or inspect the unmodified source before applying the text replacements.");
        }

        return string.Equals(sourceFileId, "default", StringComparison.OrdinalIgnoreCase)
            ? SubmitTemplateAsync<object>(caller, sourceFileId, JobKind.RenderPreview, payload: null, cancellationToken)
            : SubmitAsync<object>(caller, sourceFileId, JobKind.RenderPreview, payload: null, cancellationToken);
    }

    private static string? CreateAnalyzeSubmissionKey(
        CallerContext caller,
        string sourceFileId,
        bool includeLayouts) => string.IsNullOrWhiteSpace(caller.MessageId)
        ? null
        : string.Join(
            ':',
            caller.UserScope,
            caller.ConversationScope,
            caller.MessageId,
            sourceFileId.Trim().ToLowerInvariant(),
            includeLayouts);

    public async Task<JobReceipt> SubmitReplaceTextAsync(
        CallerContext caller,
        string sourceFileId,
        IReadOnlyList<TextReplacement> replacements,
        CancellationToken cancellationToken,
        string? previousJobId = null,
        bool isFinalBatch = true)
    {
        var payload = new ReplaceTextJobPayload(replacements, isFinalBatch);
        var workflowKey = CreateMessageSourceKey(caller, sourceFileId);
        var now = timeProvider.GetUtcNow();
        RemoveExpiredTextEditWorkflows(now);
        if (string.IsNullOrWhiteSpace(previousJobId))
        {
            if (workflowKey is not null
                && textEditWorkflows.TryGetValue(workflowKey, out var existingWorkflow))
            {
                var existingJob = await repository.GetAsync(
                    existingWorkflow.LatestJobId,
                    cancellationToken).ConfigureAwait(false);
                if (existingJob is null || existingJob.State is JobState.Failed or JobState.Canceled)
                {
                    textEditWorkflows.TryRemove(
                        new KeyValuePair<string, TextEditWorkflow>(workflowKey, existingWorkflow));
                }
                else
                {
                    throw new PptxValidationException(
                        "text_edit_workflow_already_started",
                        $"Text editing already started for this user message. Continue from latest job_id {existingWorkflow.LatestJobId}; final_batch_submitted={existingWorkflow.FinalBatchSubmitted.ToString().ToLowerInvariant()}. Do not restart from the source presentation.");
                }
            }

            var receipt = await SubmitAsync(
                caller,
                sourceFileId,
                JobKind.ReplaceText,
                payload,
                cancellationToken).ConfigureAwait(false);
            if (workflowKey is not null)
            {
                textEditWorkflows[workflowKey] = new TextEditWorkflow(
                    receipt.JobId,
                    isFinalBatch,
                    now.AddHours(24));
            }

            return receipt;
        }

        if (workflowKey is not null
            && textEditWorkflows.TryGetValue(workflowKey, out var currentWorkflow)
            && !string.Equals(currentWorkflow.LatestJobId, previousJobId, StringComparison.Ordinal))
        {
            var currentJob = await repository.GetAsync(
                currentWorkflow.LatestJobId,
                cancellationToken).ConfigureAwait(false);
            if (currentJob is null || currentJob.State is JobState.Failed or JobState.Canceled)
            {
                textEditWorkflows.TryRemove(
                    new KeyValuePair<string, TextEditWorkflow>(workflowKey, currentWorkflow));
            }
            else
            {
                throw new PptxValidationException(
                    "text_edit_job_superseded",
                    $"Text-edit job {previousJobId} is not the latest batch for this user message. Continue from latest job_id {currentWorkflow.LatestJobId}; final_batch_submitted={currentWorkflow.FinalBatchSubmitted.ToString().ToLowerInvariant()}.");
            }
        }

        var sourceJob = await GetOwnedAsync(caller, previousJobId, cancellationToken).ConfigureAwait(false);
        if (sourceJob.Kind != JobKind.ReplaceText || sourceJob.State != JobState.Succeeded)
        {
            throw new PptxValidationException(
                "text_edit_job_not_chainable",
                "Only a successful pptx_replace_text job can be used as the previous text-edit batch.");
        }

        var sourcePath = Path.Combine(repository.GetJobDirectory(sourceJob.Id), "presentation.pptx");
        if (!File.Exists(sourcePath))
        {
            throw new PptxValidationException(
                "source_expired",
                "The previous text-edit result is no longer available.");
        }

        var nextReceipt = await SubmitFromPathAsync(
            caller,
            sourceJob.SourceFileId,
            sourcePath,
            JobKind.ReplaceText,
            payload,
            cancellationToken).ConfigureAwait(false);
        if (workflowKey is not null)
        {
            textEditWorkflows[workflowKey] = new TextEditWorkflow(
                nextReceipt.JobId,
                isFinalBatch,
                now.AddHours(24));
        }

        return nextReceipt;
    }

    private static string? CreateMessageSourceKey(CallerContext caller, string sourceFileId) =>
        string.IsNullOrWhiteSpace(caller.MessageId)
            ? null
            : string.Join(
                ':',
                caller.UserScope,
                caller.ConversationScope,
                caller.MessageId,
                sourceFileId.Trim().ToLowerInvariant());

    private void RemoveExpiredTextEditWorkflows(DateTimeOffset now)
    {
        foreach (var entry in textEditWorkflows)
        {
            if (entry.Value.ExpiresAt <= now)
            {
                textEditWorkflows.TryRemove(entry);
            }
        }
    }

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
        SubmitTemplateAsync(caller, sourceFileId, JobKind.CreateDeck, slides, cancellationToken);

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

    public async Task<JobReceipt> SubmitVisualDeckAsync(
        CallerContext caller,
        VisualDeckSpec deck,
        bool useDefaultTemplate,
        CancellationToken cancellationToken)
    {
        VisualDeckValidator.Validate(deck, options.MaxSlides);
        ValidateImageAssetsOwned(caller, deck);
        if (!useDefaultTemplate || !templates.HasDefault)
        {
            return await SubmitGeneratedAsync(
                caller,
                JobKind.CreateVisualDeck,
                deck,
                cancellationToken,
                new VisualJobSubmission(IsRoot: true))
                .ConfigureAwait(false);
        }

        return await SubmitTemplateAsync(
            caller,
            "default",
            JobKind.CreateBrandedVisualDeck,
            new BrandedVisualDeckSpec(deck, "auto"),
            cancellationToken,
            new VisualJobSubmission(IsRoot: true)).ConfigureAwait(false);
    }

    public Task<JobReceipt> SubmitBrandedVisualDeckAsync(
        CallerContext caller,
        string sourceFileId,
        VisualDeckSpec deck,
        string templateLayoutId,
        CancellationToken cancellationToken)
    {
        VisualDeckValidator.Validate(deck, options.MaxSlides);
        ValidateImageAssetsOwned(caller, deck);
        if (string.IsNullOrWhiteSpace(templateLayoutId) || templateLayoutId.Length > 512)
        {
            throw new PptxValidationException(
                "invalid_template_layout",
                "Specify 'auto' or an exact blank template layout_id returned by pptx_analyze.");
        }

        return SubmitTemplateAsync(
            caller,
            sourceFileId,
            JobKind.CreateBrandedVisualDeck,
            new BrandedVisualDeckSpec(deck, templateLayoutId),
            cancellationToken,
            new VisualJobSubmission(IsRoot: true));
    }

    public async Task<VisualDeckStartDecision> AuthorizeVisualDeckStartAsync(
        CallerContext caller,
        bool userRequestedNewWorkflow,
        CancellationToken cancellationToken)
    {
        var visualJobs = new List<JobRecord>();
        await foreach (var candidate in repository.ListAsync(cancellationToken).ConfigureAwait(false))
        {
            if (candidate.Kind is (JobKind.CreateVisualDeck or JobKind.CreateBrandedVisualDeck)
                && string.Equals(candidate.UserScope, caller.UserScope, StringComparison.Ordinal)
                && string.Equals(candidate.ConversationScope, caller.ConversationScope, StringComparison.Ordinal))
            {
                visualJobs.Add(candidate);
            }
        }

        if (visualJobs.Any(static job => job.State is JobState.Queued or JobState.Running))
        {
            throw new PptxValidationException(
                "visual_deck_generation_in_progress",
                "A visual deck job is already queued or running in this conversation. Wait for it instead of starting the deck again.");
        }

        var rootAttempts = visualJobs
            .Where(static job => job.ParentJobId is null)
            .OrderByDescending(static job => job.CreatedAt)
            .ThenByDescending(static job => job.Id, StringComparer.Ordinal)
            .ToArray();
        var latestRootAttempt = rootAttempts.FirstOrDefault();
        if (latestRootAttempt is null)
        {
            return new VisualDeckStartDecision(AllowSubmittedReplacement: false, IsRecoveryRestart: false);
        }

        if (latestRootAttempt.State == JobState.Succeeded)
        {
            if (!userRequestedNewWorkflow)
            {
                throw new PptxValidationException(
                    "visual_deck_already_completed",
                    "A visual deck has already succeeded in this conversation. Use page-level refinement or slide insertion; do not start the whole deck again.");
            }

            return new VisualDeckStartDecision(AllowSubmittedReplacement: true, IsRecoveryRestart: false);
        }

        if (!userRequestedNewWorkflow && IsRecoverableAuthoredVisualFailure(latestRootAttempt))
        {
            throw new PptxValidationException(
                "visual_deck_failed_page_refinement_required",
                $"Visual deck job {latestRootAttempt.Id} failed while validating model-authored HTML/CSS. Replace only the page identified by its error with pptx_refine_visual_slide using this exact job ID; do not create another Design Brief or start another draft.");
        }

        var failedRootAttempts = rootAttempts
            .TakeWhile(static job => job.State is JobState.Failed or JobState.Canceled)
            .Count();
        if (failedRootAttempts >= 2)
        {
            throw new PptxValidationException(
                "visual_deck_recovery_limit_reached",
                "The initial visual deck generation and its single recovery attempt both failed. Stop and report the latest job error instead of starting again.");
        }

        return new VisualDeckStartDecision(
            AllowSubmittedReplacement: failedRootAttempts == 1,
            IsRecoveryRestart: failedRootAttempts == 1);
    }

    public async Task<JobReceipt> SubmitRefineVisualDeckAsync(
        CallerContext caller,
        string jobId,
        IReadOnlyList<VisualSlideRevision> revisions,
        CancellationToken cancellationToken)
    {
        var sourceJob = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        var mutationLock = GetVisualMutationLock(sourceJob);
        await mutationLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await SubmitRefineVisualDeckCoreAsync(
                caller,
                sourceJob,
                revisions,
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            mutationLock.Release();
        }
    }

    private async Task<JobReceipt> SubmitRefineVisualDeckCoreAsync(
        CallerContext caller,
        JobRecord sourceJob,
        IReadOnlyList<VisualSlideRevision> revisions,
        CancellationToken cancellationToken)
    {
        if ((sourceJob.State != JobState.Succeeded && !IsRecoverableAuthoredVisualFailure(sourceJob))
            || sourceJob.Kind is not (JobKind.CreateVisualDeck or JobKind.CreateBrandedVisualDeck))
        {
            throw new PptxValidationException(
                "visual_deck_job_not_refinable",
                "Only a successful visual deck, or a failed visual deck with a recoverable model-authored HTML/CSS validation error, can be refined.");
        }

        if (revisions is null || revisions.Count != 1)
        {
            throw new PptxValidationException(
                "visual_refinement_must_be_single_slide",
                "Refine exactly one problem slide per call. Do not resend multiple slides or the complete deck.");
        }

        var visualSubmission = await PrepareVisualRefinementAsync(
            caller,
            sourceJob,
            revisions[0].SlideNumber,
            cancellationToken).ConfigureAwait(false);

        if (sourceJob.Kind == JobKind.CreateBrandedVisualDeck)
        {
            var branded = sourceJob.Payload?.Deserialize<BrandedVisualDeckSpec>(SerializerOptions)
                ?? throw new PptxValidationException("invalid_job_payload", "The source branded visual deck specification is missing.");
            var refinedBrandedDeck = ApplyVisualDeckRevisions(branded.Deck, revisions, options.MaxSlides);
            VisualDeckValidator.Validate(refinedBrandedDeck, options.MaxSlides);
            ValidateImageAssetsOwned(caller, refinedBrandedDeck);
            var sourcePath = Path.Combine(repository.GetJobDirectory(sourceJob.Id), "source.pptx");
            if (!File.Exists(sourcePath))
            {
                throw new PptxValidationException("source_expired", "The source template is no longer available.");
            }

            return await SubmitFromPathAsync(
                caller,
                sourceJob.SourceFileId,
                sourcePath,
                JobKind.CreateBrandedVisualDeck,
                branded with { Deck = refinedBrandedDeck },
                cancellationToken,
                visualSubmission).ConfigureAwait(false);
        }

        var originalDeck = sourceJob.Payload?.Deserialize<VisualDeckSpec>(SerializerOptions)
            ?? throw new PptxValidationException("invalid_job_payload", "The source visual deck specification is missing.");
        var refinedDeck = ApplyVisualDeckRevisions(originalDeck, revisions, options.MaxSlides);
        VisualDeckValidator.Validate(refinedDeck, options.MaxSlides);
        ValidateImageAssetsOwned(caller, refinedDeck);
        return await SubmitGeneratedAsync(
            caller,
            JobKind.CreateVisualDeck,
            refinedDeck,
            cancellationToken,
            visualSubmission).ConfigureAwait(false);
    }

    public async Task<JobReceipt> SubmitInsertVisualSlidesAsync(
        CallerContext caller,
        string jobId,
        IReadOnlyList<VisualSlideSpec> insertedSlides,
        int? afterSlideNumber,
        CancellationToken cancellationToken)
    {
        var sourceJob = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        var mutationLock = GetVisualMutationLock(sourceJob);
        await mutationLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await SubmitInsertVisualSlidesCoreAsync(
                caller,
                sourceJob,
                insertedSlides,
                afterSlideNumber,
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            mutationLock.Release();
        }
    }

    private async Task<JobReceipt> SubmitInsertVisualSlidesCoreAsync(
        CallerContext caller,
        JobRecord sourceJob,
        IReadOnlyList<VisualSlideSpec> insertedSlides,
        int? afterSlideNumber,
        CancellationToken cancellationToken)
    {
        if (sourceJob.State != JobState.Succeeded
            || sourceJob.Kind is not (JobKind.CreateVisualDeck or JobKind.CreateBrandedVisualDeck))
        {
            throw new PptxValidationException(
                "visual_deck_job_not_insertable",
                "Only a successful visual or branded visual deck job can receive inserted slides.");
        }

        await EnsureLatestVisualJobAsync(caller, sourceJob, cancellationToken).ConfigureAwait(false);
        var insertionSubmission = new VisualJobSubmission(
            RootJobId: GetVisualRootJobId(sourceJob),
            ParentJobId: sourceJob.Id,
            RevisionRound: 0,
            RevisedSlidesInRound: []);

        if (sourceJob.Kind == JobKind.CreateBrandedVisualDeck)
        {
            var branded = sourceJob.Payload?.Deserialize<BrandedVisualDeckSpec>(SerializerOptions)
                ?? throw new PptxValidationException("invalid_job_payload", "The source branded visual deck specification is missing.");
            var extendedBrandedDeck = InsertVisualSlides(
                branded.Deck,
                insertedSlides,
                afterSlideNumber,
                options.MaxSlides);
            VisualDeckValidator.Validate(extendedBrandedDeck, options.MaxSlides);
            ValidateImageAssetsOwned(caller, extendedBrandedDeck);
            var sourcePath = Path.Combine(repository.GetJobDirectory(sourceJob.Id), "source.pptx");
            if (!File.Exists(sourcePath))
            {
                throw new PptxValidationException("source_expired", "The source template is no longer available.");
            }

            return await SubmitFromPathAsync(
                caller,
                sourceJob.SourceFileId,
                sourcePath,
                JobKind.CreateBrandedVisualDeck,
                branded with { Deck = extendedBrandedDeck },
                cancellationToken,
                insertionSubmission).ConfigureAwait(false);
        }

        var originalDeck = sourceJob.Payload?.Deserialize<VisualDeckSpec>(SerializerOptions)
            ?? throw new PptxValidationException("invalid_job_payload", "The source visual deck specification is missing.");
        var extendedDeck = InsertVisualSlides(
            originalDeck,
            insertedSlides,
            afterSlideNumber,
            options.MaxSlides);
        VisualDeckValidator.Validate(extendedDeck, options.MaxSlides);
        ValidateImageAssetsOwned(caller, extendedDeck);
        return await SubmitGeneratedAsync(
            caller,
            JobKind.CreateVisualDeck,
            extendedDeck,
            cancellationToken,
            insertionSubmission).ConfigureAwait(false);
    }

    public async Task<string> GetLatestSuccessfulVisualJobIdAsync(
        CallerContext caller,
        CancellationToken cancellationToken)
    {
        JobRecord? latest = null;
        await foreach (var candidate in repository.ListAsync(cancellationToken).ConfigureAwait(false))
        {
            if (candidate.State != JobState.Succeeded
                || candidate.Kind is not (JobKind.CreateVisualDeck or JobKind.CreateBrandedVisualDeck)
                || !string.Equals(candidate.UserScope, caller.UserScope, StringComparison.Ordinal)
                || !string.Equals(candidate.ConversationScope, caller.ConversationScope, StringComparison.Ordinal))
            {
                continue;
            }

            if (latest is null || candidate.CreatedAt > latest.CreatedAt)
            {
                latest = candidate;
            }
        }

        return latest?.Id
            ?? throw new PptxValidationException(
                "visual_job_not_found",
                "No successful visual deck job was found in this conversation. Create the deck before refining a slide.");
    }

    public async Task<JobView> GetAsync(CallerContext caller, string jobId, CancellationToken cancellationToken)
    {
        var job = string.Equals(jobId, "latest", StringComparison.OrdinalIgnoreCase)
            ? await GetLatestOwnedAsync(caller, cancellationToken).ConfigureAwait(false)
            : await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
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
            job.ErrorMessage,
            job.VisualRootJobId,
            job.VisualRevisionRound,
            job.VisualRevisedSlidesInRound);
    }

    public async Task<JobView?> WaitForTerminalAsync(
        CallerContext caller,
        string jobId,
        TimeSpan maximumWait,
        CancellationToken cancellationToken)
    {
        var current = await GetAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        if (current.Status is JobState.Succeeded or JobState.Failed or JobState.Canceled)
        {
            return current;
        }

        var resolvedJobId = current.JobId;
        var startedAt = timeProvider.GetTimestamp();
        while (timeProvider.GetElapsedTime(startedAt) < maximumWait)
        {
            await Task.Delay(TimeSpan.FromMilliseconds(250), cancellationToken).ConfigureAwait(false);
            current = await GetAsync(caller, resolvedJobId, cancellationToken).ConfigureAwait(false);
            if (current.Status is JobState.Succeeded or JobState.Failed or JobState.Canceled)
            {
                return current;
            }
        }

        return null;
    }

    public async Task<JobView> WaitAsync(
        CallerContext caller,
        string jobId,
        TimeSpan maximumWait,
        CancellationToken cancellationToken)
    {
        var initial = await GetAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        if (initial.Status is JobState.Succeeded or JobState.Failed or JobState.Canceled)
        {
            return initial;
        }

        var terminal = await WaitForTerminalAsync(
            caller,
            initial.JobId,
            maximumWait,
            cancellationToken).ConfigureAwait(false);
        return terminal
            ?? await GetAsync(caller, initial.JobId, cancellationToken).ConfigureAwait(false);
    }

    private async Task<JobRecord> GetLatestOwnedAsync(
        CallerContext caller,
        CancellationToken cancellationToken)
    {
        JobRecord? latest = null;
        await foreach (var candidate in repository.ListAsync(cancellationToken).ConfigureAwait(false))
        {
            if (!string.Equals(candidate.UserScope, caller.UserScope, StringComparison.Ordinal)
                || !string.Equals(candidate.ConversationScope, caller.ConversationScope, StringComparison.Ordinal))
            {
                continue;
            }

            if (latest is null
                || candidate.CreatedAt > latest.CreatedAt
                || (candidate.CreatedAt == latest.CreatedAt
                    && string.CompareOrdinal(candidate.Id, latest.Id) > 0))
            {
                latest = candidate;
            }
        }

        return latest
            ?? throw new PptxValidationException(
                "job_not_found",
                "No PowerPoint job was found in this conversation.");
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
        var job = await GetOwnedAsync(caller, jobId, cancellationToken).ConfigureAwait(false);
        if (job.State != JobState.Succeeded)
        {
            throw new PptxValidationException(
                "job_not_ready",
                "The job must succeed before preview images can be reviewed.");
        }

        var jobDirectory = repository.GetJobDirectory(job.Id);
        var availableSlideNumbers = job.Artifacts
            .Where(static artifact => string.Equals(artifact.MediaType, "image/png", StringComparison.Ordinal))
            .Select(static artifact => TryGetPreviewSlideNumber(artifact.FileName))
            .Where(static slideNumber => slideNumber.HasValue)
            .Select(static slideNumber => slideNumber!.Value)
            .Order()
            .ToArray();
        ValidatePreviewSelection(slideNumbers, availableSlideNumbers);

        var images = new List<PreviewImageData>(slideNumbers.Count);
        foreach (var slideNumber in slideNumbers)
        {
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

    internal static void ValidatePreviewSelection(
        IReadOnlyList<int>? requestedSlideNumbers,
        IReadOnlyList<int> availableSlideNumbers)
    {
        if (requestedSlideNumbers is null
            || requestedSlideNumbers.Count is < 1 or > 4
            || requestedSlideNumbers.Distinct().Count() != requestedSlideNumbers.Count)
        {
            throw new PptxValidationException(
                "preview_selection_invalid",
                "Select between 1 and 4 distinct slide numbers per visual review call.");
        }

        if (availableSlideNumbers.Count == 0)
        {
            throw new PptxValidationException(
                "preview_slide_not_found",
                "This job has no preview slides.");
        }

        var available = availableSlideNumbers.ToHashSet();
        var invalid = requestedSlideNumbers.Where(slideNumber => !available.Contains(slideNumber)).ToArray();
        if (invalid.Length > 0)
        {
            throw new PptxValidationException(
                "preview_slide_invalid",
                $"This job has {availableSlideNumbers.Count} slides; valid slide numbers are " +
                $"{availableSlideNumbers[0]} through {availableSlideNumbers[^1]}. " +
                $"Invalid requested slide numbers: {string.Join(", ", invalid)}.");
        }
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

    private void ValidateImageAssetsOwned(CallerContext caller, VisualDeckSpec deck)
    {
        foreach (var slide in deck.Slides)
        {
            if (slide.Media is null)
            {
                continue;
            }

            if (imageAssets is null)
            {
                throw new PptxValidationException(
                    "visual_media_unavailable",
                    "Image asset resolution is unavailable in this server configuration.");
            }

            imageAssets.GetOwned(caller, slide.Media.AssetId);
        }
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

        var materializedRevisions = MaterializeVisualObjectReferences(originalDeck, revisions);
        materializedRevisions = MaterializeSpeakerNotes(originalDeck, materializedRevisions);
        ValidateBrandProfileRevisions(originalDeck, materializedRevisions);
        var revisionsBySlide = materializedRevisions.ToDictionary(static revision => revision.SlideNumber);
        var slides = originalDeck.Slides
            .Select((slide, index) => revisionsBySlide.TryGetValue(index + 1, out var revision)
                ? revision.Slide
                : slide)
            .ToArray();
        return originalDeck with { Slides = slides };
    }

    private static VisualSlideRevision[] MaterializeVisualObjectReferences(
        VisualDeckSpec originalDeck,
        IReadOnlyList<VisualSlideRevision> revisions)
    {
        var materialized = new VisualSlideRevision[revisions.Count];
        for (var index = 0; index < revisions.Count; index++)
        {
            var revision = revisions[index];
            var originalReferences = originalDeck.Slides[revision.SlideNumber - 1].VisualObjects ?? [];
            var replacementReferences = revision.Slide.VisualObjects;

            if (replacementReferences is not { Count: > 0 })
            {
                materialized[index] = originalReferences.Count == 0
                    ? revision
                    : revision with
                    {
                        Slide = revision.Slide with { VisualObjects = originalReferences },
                    };
                continue;
            }

            if (!replacementReferences.Select(static item => item.AssetId)
                    .SequenceEqual(originalReferences.Select(static item => item.AssetId), StringComparer.Ordinal))
            {
                throw new PptxValidationException(
                    "visual_object_binding_mismatch",
                    $"Slide {revision.SlideNumber} must preserve its prepared visual object asset IDs during refinement. Omit visualObjects to inherit the stored list.");
            }

            materialized[index] = revision;
        }

        return materialized;
    }

    private static VisualSlideRevision[] MaterializeSpeakerNotes(
        VisualDeckSpec originalDeck,
        IReadOnlyList<VisualSlideRevision> revisions) =>
        revisions
            .Select(revision => revision.Slide.SpeakerNotes is not null
                ? revision
                : revision with
                {
                    Slide = revision.Slide with
                    {
                        SpeakerNotes = originalDeck.Slides[revision.SlideNumber - 1].SpeakerNotes,
                    },
                })
            .ToArray();

    internal static VisualDeckSpec InsertVisualSlides(
        VisualDeckSpec originalDeck,
        IReadOnlyList<VisualSlideSpec> insertedSlides,
        int? afterSlideNumber,
        int maximumSlides)
    {
        if (originalDeck is null
            || originalDeck.Slides is null
            || originalDeck.Slides.Count is < 1
            || originalDeck.Slides.Count > maximumSlides
            || insertedSlides is null
            || insertedSlides.Count < 1
            || insertedSlides.Any(static slide => slide is null)
            || originalDeck.Slides.Count + insertedSlides.Count > maximumSlides)
        {
            throw new PptxValidationException(
                "visual_deck_insert_invalid",
                $"Provide 1 or more visual slides while keeping the combined deck within {maximumSlides} slides.");
        }

        if (originalDeck.BrandProfileBinding is not null)
        {
            throw new PptxValidationException(
                "brand_profile_insert_requires_design_brief",
                "Slide insertion into a Brand Profile-bound deck is unavailable in phase 1 because the new slides have no validated Asset Plan or immutable recipe binding. Create a separately validated workflow when insertion support is added; do not bypass the existing profile contract.");
        }

        var insertionIndex = afterSlideNumber ?? originalDeck.Slides.Count;
        if (insertionIndex < 0 || insertionIndex > originalDeck.Slides.Count)
        {
            throw new PptxValidationException(
                "visual_deck_insert_position_invalid",
                $"afterSlideNumber must be between 0 and {originalDeck.Slides.Count}; omit it to append.");
        }

        var combinedSlides = originalDeck.Slides
            .Take(insertionIndex)
            .Concat(insertedSlides)
            .Concat(originalDeck.Slides.Skip(insertionIndex))
            .ToArray();
        return originalDeck with { Slides = combinedSlides };
    }

    private static void ValidateBrandProfileRevisions(
        VisualDeckSpec originalDeck,
        IReadOnlyList<VisualSlideRevision> revisions)
    {
        var brandBinding = originalDeck.BrandProfileBinding;
        if (brandBinding is null)
        {
            return;
        }

        if (brandBinding.Profile is null
            || brandBinding.Slides is null
            || brandBinding.Slides.Count != originalDeck.Slides.Count
            || brandBinding.Slides.Select(static slide => slide.SlideNumber).Distinct().Count()
                != brandBinding.Slides.Count)
        {
            throw new PptxValidationException(
                "invalid_job_payload",
                "The stored Brand Profile recipe binding is incomplete or invalid.");
        }

        if (brandBinding.DesignBriefAudit is { } audit
            && (audit.Assumptions is null
                || audit.Assumptions.Any(static assumption =>
                    assumption is null || assumption.Status == DesignAssumptionStatus.NeedsConfirmation)
                || audit.Slides is null
                || audit.Slides.Count != originalDeck.Slides.Count
                || audit.Slides.Select(static slide => slide.SlideNumber).Distinct().Count()
                    != audit.Slides.Count
                || audit.Slides.Any(slide =>
                    slide is null || slide.SlideNumber < 1 || slide.SlideNumber > originalDeck.Slides.Count)))
        {
            throw new PptxValidationException(
                "invalid_job_payload",
                "The stored Design Brief audit snapshot is incomplete or invalid.");
        }

        foreach (var revision in revisions)
        {
            var recipe = brandBinding.Slides.SingleOrDefault(binding =>
                binding.SlideNumber == revision.SlideNumber);
            if (recipe is null)
            {
                throw new PptxValidationException(
                    "invalid_job_payload",
                    $"The stored Brand Profile recipe binding is missing slide {revision.SlideNumber}.");
            }

            var slide = revision.Slide;
            if (!string.Equals(slide.RecipeId, recipe.RecipeId, StringComparison.Ordinal))
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_mismatch",
                    $"Slide {revision.SlideNumber}.recipeId must remain {recipe.RecipeId} during refinement.");
            }

            if (slide.Kind != recipe.SemanticKind)
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_kind_mismatch",
                    $"Slide {revision.SlideNumber}.kind must remain {recipe.SemanticKind} during refinement.");
            }

            var effectiveDensity = slide.Density
                ?? originalDeck.Design?.Density
                ?? "balanced";
            if (!string.Equals(effectiveDensity, recipe.Density, StringComparison.OrdinalIgnoreCase))
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_density_mismatch",
                    $"Slide {revision.SlideNumber} effective density must remain {recipe.Density} during refinement.");
            }

            if (!string.Equals(slide.Variant, recipe.Variant, StringComparison.OrdinalIgnoreCase))
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_variant_mismatch",
                    $"Slide {revision.SlideNumber}.variant must remain {recipe.Variant} during refinement.");
            }

            var assetAudit = brandBinding.DesignBriefAudit?.Slides.SingleOrDefault(item =>
                item.SlideNumber == revision.SlideNumber);
            if (assetAudit?.AssetId is { } assetId
                && (slide.Media is null
                    || !string.Equals(slide.Media.AssetId, assetId, StringComparison.Ordinal)
                    || !string.Equals(slide.Media.CropIntent, assetAudit.CropIntent, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(slide.Media.TextPosition, assetAudit.TextSafeArea, StringComparison.OrdinalIgnoreCase)))
            {
                throw new PptxValidationException(
                    "visual_media_asset_binding_mismatch",
                    $"Slide {revision.SlideNumber} must preserve the verified image asset, crop intent, and text position during refinement.");
            }
        }
    }

    private async Task<VisualJobSubmission> PrepareVisualRefinementAsync(
        CallerContext caller,
        JobRecord sourceJob,
        int slideNumber,
        CancellationToken cancellationToken)
    {
        await EnsureLatestVisualJobAsync(caller, sourceJob, cancellationToken).ConfigureAwait(false);

        var sourceRound = sourceJob.VisualRevisionRound;
        var revisedInSourceRound = sourceJob.VisualRevisedSlidesInRound ?? [];
        var nextRound = sourceRound == 0
            ? 1
            : revisedInSourceRound.Contains(slideNumber)
                ? sourceRound + 1
                : sourceRound;
        if (nextRound > MaximumVisualRevisionRounds)
        {
            throw new PptxValidationException(
                "visual_refinement_limit_reached",
                $"Slide {slideNumber} has already been refined in all {MaximumVisualRevisionRounds} allowed visual review rounds. Stop refining and return the latest successful deck.");
        }

        var revisedSlides = nextRound == sourceRound
            ? revisedInSourceRound.Append(slideNumber).Distinct().Order().ToArray()
            : [slideNumber];
        return new VisualJobSubmission(
            RootJobId: GetVisualRootJobId(sourceJob),
            ParentJobId: sourceJob.Id,
            RevisionRound: nextRound,
            RevisedSlidesInRound: revisedSlides);
    }

    private async Task EnsureLatestVisualJobAsync(
        CallerContext caller,
        JobRecord sourceJob,
        CancellationToken cancellationToken)
    {
        var rootJobId = GetVisualRootJobId(sourceJob);
        JobRecord? latestSucceeded = null;
        await foreach (var candidate in repository.ListAsync(cancellationToken).ConfigureAwait(false))
        {
            if (candidate.Kind is not (JobKind.CreateVisualDeck or JobKind.CreateBrandedVisualDeck)
                || !string.Equals(candidate.UserScope, caller.UserScope, StringComparison.Ordinal)
                || !string.Equals(candidate.ConversationScope, caller.ConversationScope, StringComparison.Ordinal)
                || !string.Equals(GetVisualRootJobId(candidate), rootJobId, StringComparison.Ordinal))
            {
                continue;
            }

            if (candidate.Id != sourceJob.Id && candidate.State is JobState.Queued or JobState.Running)
            {
                throw new PptxValidationException(
                    "visual_deck_operation_in_progress",
                    "Another operation for this visual deck is already queued or running. Wait for it instead of branching from an older result.");
            }

            if (candidate.State == JobState.Succeeded
                && (latestSucceeded is null
                    || candidate.CreatedAt > latestSucceeded.CreatedAt
                    || (candidate.CreatedAt == latestSucceeded.CreatedAt
                        && string.CompareOrdinal(candidate.Id, latestSucceeded.Id) > 0)))
            {
                latestSucceeded = candidate;
            }
        }

        if (latestSucceeded is not null && latestSucceeded.Id != sourceJob.Id)
        {
            throw new PptxValidationException(
                "visual_deck_job_superseded",
                $"Job {sourceJob.Id} is not the latest successful version of this deck. Use jobId=latest so prior page-level changes are preserved.");
        }
    }

    private static string GetVisualRootJobId(JobRecord job) =>
        string.IsNullOrWhiteSpace(job.VisualRootJobId) ? job.Id : job.VisualRootJobId;

    internal static bool IsRecoverableAuthoredVisualFailure(JobRecord job) =>
        IsRecoverableAuthoredVisualFailure(job.State, job.ErrorCode, job.ErrorMessage);

    internal static bool IsRecoverableAuthoredVisualFailure(
        JobState state,
        string? errorCode,
        string? errorMessage) =>
        state == JobState.Failed
        && (string.Equals(errorCode, "visual_authored_html_invalid", StringComparison.Ordinal)
            || (string.Equals(errorCode, "visual_renderer_failed", StringComparison.Ordinal)
                && errorMessage is { } message
                && (message.Contains("Model-authored", StringComparison.OrdinalIgnoreCase)
                    || message.Contains("Default-template cover", StringComparison.OrdinalIgnoreCase)
                    || message.Contains("Default-template body", StringComparison.OrdinalIgnoreCase))));

    private SemaphoreSlim GetVisualMutationLock(JobRecord job)
    {
        var hash = (uint)StringComparer.Ordinal.GetHashCode(GetVisualRootJobId(job));
        return visualMutationLocks[hash % visualMutationLocks.Length];
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

    private async Task<JobReceipt> SubmitTemplateAsync<TPayload>(
        CallerContext caller,
        string sourceFileId,
        JobKind kind,
        TPayload? payload,
        CancellationToken cancellationToken,
        VisualJobSubmission? visualSubmission = null)
    {
        var input = string.Equals(sourceFileId, "default", StringComparison.OrdinalIgnoreCase)
            ? await templates.ResolveDefaultAsync(cancellationToken).ConfigureAwait(false)
            : await inputFileResolver.ResolveAsync(caller, sourceFileId, cancellationToken).ConfigureAwait(false);
        return await SubmitFromPathAsync(
            caller,
            input.FileId,
            input.Path,
            kind,
            payload,
            cancellationToken,
            visualSubmission).ConfigureAwait(false);
    }

    private async Task<JobReceipt> SubmitFromPathAsync<TPayload>(
        CallerContext caller,
        string sourceFileId,
        string sourcePath,
        JobKind kind,
        TPayload? payload,
        CancellationToken cancellationToken,
        VisualJobSubmission? visualSubmission = null)
    {
        var now = timeProvider.GetUtcNow();
        var jobId = Guid.NewGuid().ToString("N");
        var job = new JobRecord
        {
            Id = jobId,
            Kind = kind,
            State = JobState.Queued,
            UserScope = caller.UserScope,
            ConversationScope = caller.ConversationScope,
            SourceFileId = sourceFileId,
            CreatedAt = now,
            ExpiresAt = now.AddDays(options.RetentionDays),
            ProgressPercent = 0,
            Payload = payload is null ? null : JsonSerializer.SerializeToElement(payload, SerializerOptions),
            ParentJobId = visualSubmission?.ParentJobId,
            VisualRootJobId = visualSubmission?.IsRoot == true
                ? jobId
                : visualSubmission?.RootJobId,
            VisualRevisionRound = visualSubmission?.RevisionRound ?? 0,
            VisualRevisedSlidesInRound = visualSubmission?.RevisedSlidesInRound ?? [],
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
        CancellationToken cancellationToken,
        VisualJobSubmission? visualSubmission = null)
    {
        var now = timeProvider.GetUtcNow();
        var jobId = Guid.NewGuid().ToString("N");
        var job = new JobRecord
        {
            Id = jobId,
            Kind = kind,
            State = JobState.Queued,
            UserScope = caller.UserScope,
            ConversationScope = caller.ConversationScope,
            SourceFileId = "generated",
            CreatedAt = now,
            ExpiresAt = now.AddDays(options.RetentionDays),
            ProgressPercent = 0,
            Payload = JsonSerializer.SerializeToElement(payload, SerializerOptions),
            ParentJobId = visualSubmission?.ParentJobId,
            VisualRootJobId = visualSubmission?.IsRoot == true
                ? jobId
                : visualSubmission?.RootJobId,
            VisualRevisionRound = visualSubmission?.RevisionRound ?? 0,
            VisualRevisedSlidesInRound = visualSubmission?.RevisedSlidesInRound ?? [],
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

internal sealed record AnalyzeJobPayload(bool IncludeLayouts);

internal sealed record AnalyzeSubmission(
    Lazy<Task<JobReceipt>> Receipt,
    DateTimeOffset ExpiresAt);

internal sealed record TextEditWorkflow(
    string LatestJobId,
    bool FinalBatchSubmitted,
    DateTimeOffset ExpiresAt);

internal sealed record ReplaceTextJobPayload(
    IReadOnlyList<TextReplacement> Replacements,
    bool IsFinalBatch);

public sealed record VisualDeckStartDecision(
    bool AllowSubmittedReplacement,
    bool IsRecoveryRestart);

internal sealed record VisualJobSubmission(
    bool IsRoot = false,
    string? RootJobId = null,
    string? ParentJobId = null,
    int RevisionRound = 0,
    IReadOnlyList<int>? RevisedSlidesInRound = null);
