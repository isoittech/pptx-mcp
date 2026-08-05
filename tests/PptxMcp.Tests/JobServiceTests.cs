using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Jobs;
using PptxMcp.Domain;
using PptxMcp.Security;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class JobServiceTests
{
    [Theory]
    [InlineData("preview/slide-1.png", 1)]
    [InlineData("preview/slide-01.png", 1)]
    [InlineData("preview/slide-50.png", 50)]
    public void ParsesPaddedAndUnpaddedPreviewSlideNumbers(string fileName, int expected)
    {
        Assert.Equal(expected, JobService.TryGetPreviewSlideNumber(fileName));
    }

    [Theory]
    [InlineData("slide-01.png")]
    [InlineData("preview/slide-one.png")]
    [InlineData("preview/slide-01.jpg")]
    public void RejectsUnexpectedPreviewArtifactNames(string fileName)
    {
        Assert.Null(JobService.TryGetPreviewSlideNumber(fileName));
    }

    [Fact]
    public void AppliesOnlyRequestedDeckRevisionsAndPreservesLayouts()
    {
        var slides = new[]
        {
            new DeckSlideSpec("layout-1", [new DeckField("Original 1", ShapeId: 2)]),
            new DeckSlideSpec("layout-2", [new DeckField("Original 2", ShapeId: 3)]),
            new DeckSlideSpec("layout-3", [new DeckField("Original 3", ShapeId: 4)]),
        };
        var revisions = new[]
        {
            new DeckSlideRevision(2, [new DeckField("Revised 2", ShapeId: 3)]),
        };

        var result = JobService.ApplyDeckRevisions(slides, revisions, 50);

        Assert.Equal("Original 1", result[0].Fields[0].Text);
        Assert.Equal("layout-2", result[1].LayoutId);
        Assert.Equal("Revised 2", result[1].Fields[0].Text);
        Assert.Equal("Original 3", result[2].Fields[0].Text);
    }

    [Fact]
    public void RejectsDuplicateDeckRevisionSlideNumbers()
    {
        var slides = new[]
        {
            new DeckSlideSpec("layout-1", [new DeckField("Original", ShapeId: 2)]),
        };
        var revisions = new[]
        {
            new DeckSlideRevision(1, [new DeckField("First", ShapeId: 2)]),
            new DeckSlideRevision(1, [new DeckField("Second", ShapeId: 2)]),
        };

        var error = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyDeckRevisions(slides, revisions, 50));

        Assert.Equal("deck_revision_invalid", error.Code);
    }

    [Fact]
    public void AppliesOnlyRequestedVisualSlideRevisionsAndPreservesCreativeDirection()
    {
        var original = new VisualDeckSpec(
            "危機対応",
            [
                new VisualSlideSpec(VisualSlideKind.Title, "初動72時間"),
                new VisualSlideSpec(VisualSlideKind.Process, "初動", Steps:
                [
                    new VisualStepSpec("検知"),
                    new VisualStepSpec("封じ込め"),
                    new VisualStepSpec("復旧"),
                ]),
            ],
            Design: new VisualDesignSpec("bold", "airy", "orbit"));
        var revisions = new[]
        {
            new VisualSlideRevision(
                2,
                new VisualSlideSpec(VisualSlideKind.Statement, "判断", Body: "事業継続を最優先する")),
        };

        var result = JobService.ApplyVisualDeckRevisions(original, revisions, 50);

        Assert.Equal(VisualSlideKind.Title, result.Slides[0].Kind);
        Assert.Equal(VisualSlideKind.Statement, result.Slides[1].Kind);
        Assert.Equal("bold", result.Design?.Style);
        Assert.Equal("orbit", result.Design?.Motif);
    }

    [Fact]
    public void RejectsDuplicateVisualDeckRevisionSlideNumbers()
    {
        var original = new VisualDeckSpec(
            "危機対応",
            [new VisualSlideSpec(VisualSlideKind.Title, "初動72時間")]);
        var revisions = new[]
        {
            new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, "A")),
            new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, "B")),
        };

        var error = Assert.Throws<PptxValidationException>(() =>
            JobService.ApplyVisualDeckRevisions(original, revisions, 50));

        Assert.Equal("visual_deck_revision_invalid", error.Code);
    }

    [Fact]
    public async Task LatestJobIsResolvedWithinCallerConversationIncludingRunningJobs()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);

        try
        {
            var (service, repository) = CreateJobService(storageRoot, uploadsRoot);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var now = DateTimeOffset.UtcNow;

            await repository.CreateAsync(CreateJob(
                "11111111111111111111111111111111",
                caller,
                JobState.Succeeded,
                now.AddMinutes(-1)), CancellationToken.None);
            await repository.CreateAsync(CreateJob(
                "22222222222222222222222222222222",
                caller,
                JobState.Running,
                now), CancellationToken.None);
            await repository.CreateAsync(CreateJob(
                "33333333333333333333333333333333",
                new CallerContext("user-1", "conversation-2", null),
                JobState.Succeeded,
                now.AddMinutes(1)), CancellationToken.None);

            var result = await service.GetAsync(caller, "latest", CancellationToken.None);

            Assert.Equal("22222222222222222222222222222222", result.JobId);
            Assert.Equal(JobState.Running, result.Status);
        }
        finally
        {
            if (Directory.Exists(storageRoot))
            {
                Directory.Delete(storageRoot, recursive: true);
            }

            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task TerminalJobWaitReturnsImmediatelyForSucceededJob()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);

        try
        {
            var (service, repository) = CreateJobService(storageRoot, uploadsRoot);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var job = CreateJob(
                "44444444444444444444444444444444",
                caller,
                JobState.Succeeded,
                DateTimeOffset.UtcNow);
            await repository.CreateAsync(job, CancellationToken.None);

            var result = await service.WaitForTerminalAsync(
                caller,
                job.Id,
                TimeSpan.FromSeconds(1),
                CancellationToken.None);

            Assert.NotNull(result);
            Assert.Equal(JobState.Succeeded, result.Status);
        }
        finally
        {
            if (Directory.Exists(storageRoot))
            {
                Directory.Delete(storageRoot, recursive: true);
            }

            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    private static (JobService Service, FileJobRepository Repository) CreateJobService(
        string storageRoot,
        string uploadsRoot)
    {
        var options = Options.Create(new PptxMcpOptions
        {
            StorageRoot = storageRoot,
            LibreChatUploadsRoot = uploadsRoot,
            SigningKey = new string('k', 32),
            MaxQueueDepth = 12,
        });
        var repository = new FileJobRepository(options);
        var guard = new PptxPackageGuard(options);
        var service = new JobService(
            repository,
            new InputFileResolver(options, guard),
            guard,
            new JobChannel(options),
            new JobCancellationRegistry(),
            new ArtifactTokenService(options, TimeProvider.System),
            options,
            TimeProvider.System);
        return (service, repository);
    }

    private static JobRecord CreateJob(
        string id,
        CallerContext caller,
        JobState state,
        DateTimeOffset createdAt) =>
        new()
        {
            Id = id,
            Kind = JobKind.CreateBrandedVisualDeck,
            State = state,
            UserScope = caller.UserScope,
            ConversationScope = caller.ConversationScope,
            SourceFileId = "generated",
            CreatedAt = createdAt,
            ExpiresAt = createdAt.AddDays(7),
            ProgressPercent = state == JobState.Succeeded ? 100 : 50,
        };
}
