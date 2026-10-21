using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Security;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class SlideExportJobTests : IDisposable
{
    private static readonly CallerContext Caller = new("test-user", "test-conversation", "test-message");
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly string directory = Path.Combine(Path.GetTempPath(), $"pptx-selection-{Guid.NewGuid():N}");
    private readonly FileJobRepository repository;
    private readonly JobService service;
    private readonly string uploads;

    public SlideExportJobTests()
    {
        uploads = Path.Combine(directory, "uploads");
        Directory.CreateDirectory(uploads);
        var options = Options.Create(new PptxMcpOptions
        {
            StorageRoot = Path.Combine(directory, "jobs"),
            LibreChatUploadsRoot = uploads,
            SigningKey = new string('k', 32),
        });
        repository = new FileJobRepository(options);
        var guard = new PptxPackageGuard(options);
        service = new JobService(repository, new InputFileResolver(options, guard), new TemplateRegistry(options, guard),
            guard, new JobChannel(options), new JobCancellationRegistry(),
            new ArtifactTokenService(options, TimeProvider.System), options, TimeProvider.System);
    }

    [Fact]
    public async Task ExportCopiesTheExactCompletedArtifactWithoutOverwritingIt()
    {
        var source = await CreateSourceAsync();
        var sourcePath = Path.Combine(repository.GetJobDirectory(source.Id), "presentation.pptx");
        var bytes = await File.ReadAllBytesAsync(sourcePath);
        var receipt = await service.SubmitExportSlidesAsync(Caller, new SlideSelection([2]), null, source.Id, CancellationToken.None);
        var job = await repository.GetAsync(receipt.JobId, CancellationToken.None);
        Assert.Equal(JobKind.ExportSlides, job!.Kind);
        Assert.Equal([2], job.Payload!.Value.Deserialize<SlideSelection>(JsonOptions)!.SlideNumbers);
        Assert.Equal(bytes, await File.ReadAllBytesAsync(Path.Combine(repository.GetJobDirectory(job.Id), "source.pptx")));
        Assert.Equal(bytes, await File.ReadAllBytesAsync(sourcePath));
    }

    [Theory]
    [InlineData("other-user", "test-conversation")]
    [InlineData("test-user", "other-conversation")]
    public async Task ExportRejectsAnotherUserOrConversation(string user, string conversation)
    {
        var source = await CreateSourceAsync();
        var error = await Assert.ThrowsAsync<PptxValidationException>(() => service.SubmitExportSlidesAsync(
            new CallerContext(user, conversation, null), new SlideSelection([1]), null, source.Id, CancellationToken.None));
        Assert.Equal("job_not_found", error.Code);
    }

    [Theory]
    [InlineData(JobState.Queued)]
    [InlineData(JobState.Running)]
    [InlineData(JobState.Failed)]
    public async Task ExportRejectsAnUnfinishedSource(JobState state)
    {
        var source = await CreateSourceAsync();
        await repository.UpdateAsync(source.Id, job => job with { State = state }, CancellationToken.None);
        var error = await Assert.ThrowsAsync<PptxValidationException>(() => service.SubmitExportSlidesAsync(
            Caller, new SlideSelection([1]), null, source.Id, CancellationToken.None));
        Assert.Equal("slide_export_source_not_ready", error.Code);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportRejectsExpiredSourcesEvenBeforeCleanup(bool downloaded)
    {
        var source = await CreateSourceAsync();
        await repository.UpdateAsync(source.Id, job => downloaded
            ? job with { FirstDownloadedAt = DateTimeOffset.UtcNow.AddHours(-25) }
            : job with { ExpiresAt = DateTimeOffset.UtcNow.AddMinutes(-1) }, CancellationToken.None);
        var error = await Assert.ThrowsAsync<PptxValidationException>(() => service.SubmitExportSlidesAsync(
            Caller, new SlideSelection([1]), null, source.Id, CancellationToken.None));
        Assert.Equal("source_expired", error.Code);
    }

    [Fact]
    public async Task ExportKeepsTheCurrentAttachmentScope()
    {
        var source = TestPresentationFactory.CreateSlides("One", "Two");
        var userDirectory = Path.Combine(uploads, Caller.UserId);
        Directory.CreateDirectory(userDirectory);
        File.Move(source, Path.Combine(userDirectory, "file-1__sample.pptx"));
        var caller = Caller with { AttachmentFileIds = new HashSet<string> { "file-1" } };
        var receipt = await service.SubmitExportSlidesAsync(caller, new SlideSelection([2]), "latest", null, CancellationToken.None);
        Assert.Equal(JobKind.ExportSlides, (await repository.GetAsync(receipt.JobId, CancellationToken.None))!.Kind);
        var error = await Assert.ThrowsAsync<PptxValidationException>(() => service.SubmitExportSlidesAsync(
            caller with { AttachmentFileIds = new HashSet<string>() }, new SlideSelection([2]), "latest", null, CancellationToken.None));
        Assert.Equal("file_not_found", error.Code);
    }

    [Theory]
    [InlineData(null, null)]
    [InlineData("latest", "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")]
    public async Task ExportRequiresOneUnambiguousSource(string? fileId, string? jobId)
    {
        var error = await Assert.ThrowsAsync<PptxValidationException>(() => service.SubmitExportSlidesAsync(
            Caller, new SlideSelection([1]), fileId, jobId, CancellationToken.None));
        Assert.Equal("slide_export_source_required", error.Code);
    }

    [Fact]
    public async Task SingleSlideRefinementQueuesOnlyTheRequestedPageAndPreservesTheFullDeck()
    {
        var source = await CreateSourceAsync();
        var receipt = await service.SubmitRefineVisualDeckAsync(Caller, source.Id,
            [new VisualSlideRevision(2, new VisualSlideSpec(VisualSlideKind.Title, "Revised body"))],
            CancellationToken.None, outputOnlyRevisedSlide: true);
        var job = (await repository.GetAsync(receipt.JobId, CancellationToken.None))!;
        var deck = job.Payload!.Value.Deserialize<VisualDeckSpec>(JsonOptions)!;
        Assert.Equal("Revised body", Assert.Single(deck.Slides).Title);
        Assert.False(deck.FirstSlideIsCover);
        Assert.Equal([2], deck.SourceSlideNumbers);
        Assert.Equal(job.Id, job.VisualRootJobId);
        Assert.Equal(source.Id, job.ParentJobId);
        Assert.Equal(1, job.VisualRevisionRound);
        Assert.Equal(2, (await repository.GetAsync(source.Id, CancellationToken.None))!.Payload!.Value
            .Deserialize<VisualDeckSpec>(JsonOptions)!.Slides.Count);

        await repository.UpdateAsync(job.Id, item => item with { State = JobState.Succeeded }, CancellationToken.None);
        var next = await service.SubmitRefineVisualDeckAsync(Caller, job.Id,
            [new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, "Refined again"))], CancellationToken.None);
        var nextJob = (await repository.GetAsync(next.JobId, CancellationToken.None))!;
        Assert.Equal(2, nextJob.VisualRevisionRound);
        Assert.False(nextJob.Payload!.Value.Deserialize<VisualDeckSpec>(JsonOptions)!.FirstSlideIsCover);

        // The exported subset is a separate lineage, so the full deck remains available for later edits.
        var full = await service.SubmitRefineVisualDeckAsync(Caller, source.Id,
            [new VisualSlideRevision(2, new VisualSlideSpec(VisualSlideKind.Title, "Full deck edit"))], CancellationToken.None);
        Assert.Equal(2, (await repository.GetAsync(full.JobId, CancellationToken.None))!.Payload!.Value
            .Deserialize<VisualDeckSpec>(JsonOptions)!.Slides.Count);
    }

    private async Task<JobRecord> CreateSourceAsync()
    {
        var job = new JobRecord
        {
            Id = Guid.NewGuid().ToString("N"), Kind = JobKind.CreateVisualDeck, State = JobState.Succeeded,
            UserScope = Caller.UserScope, ConversationScope = Caller.ConversationScope, SourceFileId = "generated",
            CreatedAt = DateTimeOffset.UtcNow, ExpiresAt = DateTimeOffset.UtcNow.AddDays(7),
            Artifacts = [new ArtifactRecord("presentation.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", 1, true)],
            Payload = JsonSerializer.SerializeToElement(new VisualDeckSpec("Source",
                [new VisualSlideSpec(VisualSlideKind.Title, "Cover"), new VisualSlideSpec(VisualSlideKind.Title, "Body")]), JsonOptions),
        };
        await repository.CreateAsync(job, CancellationToken.None);
        File.Move(TestPresentationFactory.CreateSlides("Cover", "Body"),
            Path.Combine(repository.GetJobDirectory(job.Id), "presentation.pptx"));
        return job;
    }

    public void Dispose() => Directory.Delete(directory, recursive: true);
}
