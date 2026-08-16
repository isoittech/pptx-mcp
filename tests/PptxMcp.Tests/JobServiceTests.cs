using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Jobs;
using PptxMcp.Domain;
using PptxMcp.Security;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class JobServiceTests
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);

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
    public void VisualSlideRevisionPreservesSpeakerNotesWhenReplacementOmitsThem()
    {
        var originalNotes = new VisualSpeakerNotesSpec(
            "意思決定の理由を伝える。",
            "このページでは判断基準を説明します。");
        var original = new VisualDeckSpec(
            "Speaker notes",
            [new VisualSlideSpec(VisualSlideKind.Statement, "判断", Body: "初版", SpeakerNotes: originalNotes)]);

        var inherited = JobService.ApplyVisualDeckRevisions(
            original,
            [new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Statement, "判断", Body: "修正版"))],
            50);
        var replacedNotes = new VisualSpeakerNotesSpec("修正理由を伝える。", "修正後の説明です。");
        var replaced = JobService.ApplyVisualDeckRevisions(
            original,
            [
                new VisualSlideRevision(
                    1,
                    new VisualSlideSpec(
                        VisualSlideKind.Statement,
                        "判断",
                        Body: "修正版",
                        SpeakerNotes: replacedNotes)),
            ],
            50);

        Assert.Same(originalNotes, inherited.Slides[0].SpeakerNotes);
        Assert.Same(replacedNotes, replaced.Slides[0].SpeakerNotes);
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
    public void InsertsOnlyNewVisualSlidesAndPreservesCreativeDirection()
    {
        var design = new VisualDesignSpec("editorial", "airy", "ribbon");
        var original = new VisualDeckSpec(
            "経営報告",
            [
                new VisualSlideSpec(VisualSlideKind.Title, "表紙"),
                new VisualSlideSpec(VisualSlideKind.Closing, "まとめ"),
            ],
            Design: design);
        var inserted = new[]
        {
            new VisualSlideSpec(VisualSlideKind.Metrics, "追加KPI", Metrics:
            [
                new VisualMetricSpec("42", "対象拠点"),
                new VisualMetricSpec("18%", "改善率"),
            ]),
            new VisualSlideSpec(VisualSlideKind.Statement, "追加提言", Body: "重点投資を前倒しする"),
        };

        var result = JobService.InsertVisualSlides(original, inserted, 1, 50);

        Assert.Equal(["表紙", "追加KPI", "追加提言", "まとめ"], result.Slides.Select(static slide => slide.Title));
        Assert.Same(design, result.Design);
        Assert.Equal("経営報告", result.Title);
    }

    [Fact]
    public void AppendsVisualSlidesWhenPositionIsOmitted()
    {
        var original = new VisualDeckSpec(
            "経営報告",
            [new VisualSlideSpec(VisualSlideKind.Title, "表紙")]);

        var result = JobService.InsertVisualSlides(
            original,
            [new VisualSlideSpec(VisualSlideKind.Closing, "追加ページ")],
            null,
            50);

        Assert.Equal(["表紙", "追加ページ"], result.Slides.Select(static slide => slide.Title));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(3)]
    public void RejectsVisualSlideInsertionOutsideExistingDeck(int afterSlideNumber)
    {
        var original = new VisualDeckSpec(
            "経営報告",
            [
                new VisualSlideSpec(VisualSlideKind.Title, "表紙"),
                new VisualSlideSpec(VisualSlideKind.Closing, "まとめ"),
            ]);

        var error = Assert.Throws<PptxValidationException>(() =>
            JobService.InsertVisualSlides(
                original,
                [new VisualSlideSpec(VisualSlideKind.Statement, "追加", Body: "追加本文")],
                afterSlideNumber,
                50));

        Assert.Equal("visual_deck_insert_position_invalid", error.Code);
    }

    [Fact]
    public void RejectsVisualSlideInsertionBeyondMaximumDeckSize()
    {
        var original = new VisualDeckSpec(
            "経営報告",
            [
                new VisualSlideSpec(VisualSlideKind.Title, "表紙"),
                new VisualSlideSpec(VisualSlideKind.Closing, "まとめ"),
            ]);

        var error = Assert.Throws<PptxValidationException>(() =>
            JobService.InsertVisualSlides(
                original,
                [new VisualSlideSpec(VisualSlideKind.Statement, "追加", Body: "追加本文")],
                null,
                2));

        Assert.Equal("visual_deck_insert_invalid", error.Code);
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

    [Fact]
    public async Task JobWaitReturnsTerminalViewAfterBackgroundCompletion()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);

        try
        {
            var (service, repository) = CreateJobService(storageRoot, uploadsRoot);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var job = CreateJob(
                "55555555555555555555555555555555",
                caller,
                JobState.Running,
                DateTimeOffset.UtcNow);
            await repository.CreateAsync(job, CancellationToken.None);

            var wait = service.WaitAsync(
                caller,
                "latest",
                TimeSpan.FromSeconds(2),
                CancellationToken.None);
            await Task.Delay(100);
            await repository.UpdateAsync(
                job.Id,
                current => current with
                {
                    State = JobState.Succeeded,
                    ProgressPercent = 100,
                    CompletedAt = DateTimeOffset.UtcNow,
                },
                CancellationToken.None);

            var result = await wait;

            Assert.Equal(job.Id, result.JobId);
            Assert.Equal(JobState.Succeeded, result.Status);
            Assert.Equal(100, result.ProgressPercent);
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
    public async Task VisualDeckUsesConfiguredDefaultTemplateWithoutUpload()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var templatesRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-templates-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);
        Directory.CreateDirectory(templatesRoot);
        File.Move(
            TestPresentationFactory.CreateBlankBrandedTemplate(),
            Path.Combine(templatesRoot, "organization-default.pptx"));

        try
        {
            var options = Options.Create(new PptxMcpOptions
            {
                StorageRoot = storageRoot,
                LibreChatUploadsRoot = uploadsRoot,
                TemplatesRoot = templatesRoot,
                DefaultTemplateId = "organization-default",
                SigningKey = new string('k', 32),
                MaxQueueDepth = 12,
            });
            var repository = new FileJobRepository(options);
            var guard = new PptxPackageGuard(options);
            var service = new JobService(
                repository,
                new InputFileResolver(options, guard),
                new TemplateRegistry(options, guard),
                guard,
                new JobChannel(options),
                new JobCancellationRegistry(),
                new ArtifactTokenService(options, TimeProvider.System),
                options,
                TimeProvider.System);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var deck = new VisualDeckSpec(
                "既定テンプレート",
                [new VisualSlideSpec(VisualSlideKind.Title, "タイトル")]);

            var receipt = await service.SubmitVisualDeckAsync(caller, deck, true, CancellationToken.None);
            var job = await repository.GetAsync(receipt.JobId, CancellationToken.None);

            Assert.NotNull(job);
            Assert.Equal(JobKind.CreateBrandedVisualDeck, job.Kind);
            Assert.Equal("organization-default", job.SourceFileId);
            Assert.True(File.Exists(Path.Combine(repository.GetJobDirectory(job.Id), "source.pptx")));

            var analysisReceipt = await service.SubmitAnalyzeAsync(caller, "default", CancellationToken.None);
            var analysisJob = await repository.GetAsync(analysisReceipt.JobId, CancellationToken.None);

            Assert.NotNull(analysisJob);
            Assert.Equal(JobKind.Analyze, analysisJob.Kind);
            Assert.Equal("organization-default", analysisJob.SourceFileId);
        }
        finally
        {
            if (Directory.Exists(storageRoot))
            {
                Directory.Delete(storageRoot, recursive: true);
            }

            Directory.Delete(uploadsRoot, recursive: true);
            Directory.Delete(templatesRoot, recursive: true);
        }
    }

    [Fact]
    public async Task VisualDeckCanExplicitlyBypassConfiguredDefaultTemplate()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var templatesRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-templates-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);
        Directory.CreateDirectory(templatesRoot);
        File.Move(
            TestPresentationFactory.CreateBlankBrandedTemplate(),
            Path.Combine(templatesRoot, "organization-default.pptx"));

        try
        {
            var options = Options.Create(new PptxMcpOptions
            {
                StorageRoot = storageRoot,
                LibreChatUploadsRoot = uploadsRoot,
                TemplatesRoot = templatesRoot,
                DefaultTemplateId = "organization-default",
                SigningKey = new string('k', 32),
                MaxQueueDepth = 12,
            });
            var repository = new FileJobRepository(options);
            var guard = new PptxPackageGuard(options);
            var service = new JobService(
                repository,
                new InputFileResolver(options, guard),
                new TemplateRegistry(options, guard),
                guard,
                new JobChannel(options),
                new JobCancellationRegistry(),
                new ArtifactTokenService(options, TimeProvider.System),
                options,
                TimeProvider.System);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var deck = new VisualDeckSpec(
                "テンプレートなし",
                [new VisualSlideSpec(VisualSlideKind.Title, "タイトル")]);

            var receipt = await service.SubmitVisualDeckAsync(caller, deck, false, CancellationToken.None);
            var job = await repository.GetAsync(receipt.JobId, CancellationToken.None);

            Assert.NotNull(job);
            Assert.Equal(JobKind.CreateVisualDeck, job.Kind);
            Assert.Equal("generated", job.SourceFileId);
        }
        finally
        {
            if (Directory.Exists(storageRoot))
            {
                Directory.Delete(storageRoot, recursive: true);
            }

            Directory.Delete(uploadsRoot, recursive: true);
            Directory.Delete(templatesRoot, recursive: true);
        }
    }

    [Fact]
    public async Task InsertedSlidesReuseBrandedDeckTemplateAndSpecification()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var templatesRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-templates-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);
        Directory.CreateDirectory(templatesRoot);
        File.Move(
            TestPresentationFactory.CreateBlankBrandedTemplate(),
            Path.Combine(templatesRoot, "organization-default.pptx"));

        try
        {
            var options = Options.Create(new PptxMcpOptions
            {
                StorageRoot = storageRoot,
                LibreChatUploadsRoot = uploadsRoot,
                TemplatesRoot = templatesRoot,
                DefaultTemplateId = "organization-default",
                SigningKey = new string('k', 32),
                MaxQueueDepth = 12,
            });
            var repository = new FileJobRepository(options);
            var guard = new PptxPackageGuard(options);
            var service = new JobService(
                repository,
                new InputFileResolver(options, guard),
                new TemplateRegistry(options, guard),
                guard,
                new JobChannel(options),
                new JobCancellationRegistry(),
                new ArtifactTokenService(options, TimeProvider.System),
                options,
                TimeProvider.System);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var originalDeck = new VisualDeckSpec(
                "既定テンプレート",
                [new VisualSlideSpec(VisualSlideKind.Title, "表紙")],
                Design: new VisualDesignSpec("bold", "balanced", "nodes"));
            var originalReceipt = await service.SubmitVisualDeckAsync(
                caller,
                originalDeck,
                true,
                CancellationToken.None);
            await repository.UpdateAsync(
                originalReceipt.JobId,
                current => current with
                {
                    State = JobState.Succeeded,
                    ProgressPercent = 100,
                    CompletedAt = DateTimeOffset.UtcNow,
                },
                CancellationToken.None);

            var insertedReceipt = await service.SubmitInsertVisualSlidesAsync(
                caller,
                originalReceipt.JobId,
                [new VisualSlideSpec(VisualSlideKind.Statement, "追加提言", Body: "投資を前倒しする")],
                null,
                CancellationToken.None);
            var insertedJob = await repository.GetAsync(insertedReceipt.JobId, CancellationToken.None);

            Assert.NotNull(insertedJob);
            Assert.Equal(JobKind.CreateBrandedVisualDeck, insertedJob.Kind);
            Assert.Equal("organization-default", insertedJob.SourceFileId);
            Assert.True(File.Exists(Path.Combine(repository.GetJobDirectory(insertedJob.Id), "source.pptx")));
            var payload = insertedJob.Payload?.Deserialize<BrandedVisualDeckSpec>(SerializerOptions);
            Assert.NotNull(payload);
            Assert.Equal("auto", payload.TemplateLayoutId);
            Assert.Equal(["表紙", "追加提言"], payload.Deck.Slides.Select(static slide => slide.Title));
            Assert.Equal("bold", payload.Deck.Design?.Style);
        }
        finally
        {
            if (Directory.Exists(storageRoot))
            {
                Directory.Delete(storageRoot, recursive: true);
            }

            Directory.Delete(uploadsRoot, recursive: true);
            Directory.Delete(templatesRoot, recursive: true);
        }
    }

    [Fact]
    public async Task SuccessfulVisualDeckBlocksImplicitFullRestartButAllowsExplicitNewWorkflow()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);

        try
        {
            var (service, repository) = CreateJobService(storageRoot, uploadsRoot);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var receipt = await service.SubmitVisualDeckAsync(
                caller,
                new VisualDeckSpec("初版", [new VisualSlideSpec(VisualSlideKind.Title, "表紙")]),
                false,
                CancellationToken.None);
            await SetStateAsync(repository, receipt.JobId, JobState.Succeeded);

            var blocked = await Assert.ThrowsAsync<PptxValidationException>(() =>
                service.AuthorizeVisualDeckStartAsync(caller, false, CancellationToken.None));
            var explicitStart = await service.AuthorizeVisualDeckStartAsync(
                caller,
                true,
                CancellationToken.None);

            Assert.Equal("visual_deck_already_completed", blocked.Code);
            Assert.True(explicitStart.AllowSubmittedReplacement);
            Assert.False(explicitStart.IsRecoveryRestart);

            var separateDeck = await service.SubmitVisualDeckAsync(
                caller,
                new VisualDeckSpec("別資料", [new VisualSlideSpec(VisualSlideKind.Title, "別表紙")]),
                false,
                CancellationToken.None);
            await SetStateAsync(repository, separateDeck.JobId, JobState.Failed);
            var separateDeckRecovery = await service.AuthorizeVisualDeckStartAsync(
                caller,
                false,
                CancellationToken.None);

            Assert.True(separateDeckRecovery.AllowSubmittedReplacement);
            Assert.True(separateDeckRecovery.IsRecoveryRestart);
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
    public async Task InitialGenerationGetsOnlyOneRecoveryRestart()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);

        try
        {
            var (service, repository) = CreateJobService(storageRoot, uploadsRoot);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var deck = new VisualDeckSpec("復旧", [new VisualSlideSpec(VisualSlideKind.Title, "表紙")]);
            var first = await service.SubmitVisualDeckAsync(caller, deck, false, CancellationToken.None);
            await SetStateAsync(repository, first.JobId, JobState.Failed);

            var recovery = await service.AuthorizeVisualDeckStartAsync(caller, false, CancellationToken.None);
            Assert.True(recovery.AllowSubmittedReplacement);
            Assert.True(recovery.IsRecoveryRestart);

            var second = await service.SubmitVisualDeckAsync(caller, deck, false, CancellationToken.None);
            await SetStateAsync(repository, second.JobId, JobState.Failed);
            var blocked = await Assert.ThrowsAsync<PptxValidationException>(() =>
                service.AuthorizeVisualDeckStartAsync(caller, false, CancellationToken.None));

            Assert.Equal("visual_deck_recovery_limit_reached", blocked.Code);
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
    public async Task VisualRefinementIsSequentialSingleSlideAndLimitedToTwoRounds()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);

        try
        {
            var (service, repository) = CreateJobService(storageRoot, uploadsRoot);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var deck = new VisualDeckSpec(
                "差分修正",
                [
                    new VisualSlideSpec(VisualSlideKind.Title, "表紙"),
                    new VisualSlideSpec(VisualSlideKind.Statement, "要点", Body: "初版"),
                ]);
            var rootReceipt = await service.SubmitVisualDeckAsync(caller, deck, false, CancellationToken.None);
            await SetStateAsync(repository, rootReceipt.JobId, JobState.Succeeded);

            var firstSlide = await service.SubmitRefineVisualDeckAsync(
                caller,
                rootReceipt.JobId,
                [new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, "表紙1巡目"))],
                CancellationToken.None);
            await SetStateAsync(repository, firstSlide.JobId, JobState.Succeeded);
            var secondSlide = await service.SubmitRefineVisualDeckAsync(
                caller,
                firstSlide.JobId,
                [new VisualSlideRevision(2, new VisualSlideSpec(VisualSlideKind.Statement, "要点1巡目", Body: "修正"))],
                CancellationToken.None);
            await SetStateAsync(repository, secondSlide.JobId, JobState.Succeeded);
            var firstSlideAgain = await service.SubmitRefineVisualDeckAsync(
                caller,
                secondSlide.JobId,
                [new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, "表紙2巡目"))],
                CancellationToken.None);
            await SetStateAsync(repository, firstSlideAgain.JobId, JobState.Succeeded);
            var firstSlideJob = await repository.GetAsync(firstSlide.JobId, CancellationToken.None);
            var secondSlideJob = await repository.GetAsync(secondSlide.JobId, CancellationToken.None);
            var secondRoundJob = await repository.GetAsync(firstSlideAgain.JobId, CancellationToken.None);

            var thirdRound = await Assert.ThrowsAsync<PptxValidationException>(() =>
                service.SubmitRefineVisualDeckAsync(
                    caller,
                    firstSlideAgain.JobId,
                    [new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, "表紙3巡目"))],
                    CancellationToken.None));
            var staleBranch = await Assert.ThrowsAsync<PptxValidationException>(() =>
                service.SubmitRefineVisualDeckAsync(
                    caller,
                    rootReceipt.JobId,
                    [new VisualSlideRevision(2, new VisualSlideSpec(VisualSlideKind.Statement, "古い分岐", Body: "不可"))],
                    CancellationToken.None));
            var multiSlide = await Assert.ThrowsAsync<PptxValidationException>(() =>
                service.SubmitRefineVisualDeckAsync(
                    caller,
                    firstSlideAgain.JobId,
                    [
                        new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, "一括1")),
                        new VisualSlideRevision(2, new VisualSlideSpec(VisualSlideKind.Statement, "一括2", Body: "不可")),
                    ],
                    CancellationToken.None));

            Assert.Equal(1, firstSlideJob?.VisualRevisionRound);
            Assert.Equal([1], firstSlideJob?.VisualRevisedSlidesInRound);
            Assert.Equal(1, secondSlideJob?.VisualRevisionRound);
            Assert.Equal([1, 2], secondSlideJob?.VisualRevisedSlidesInRound);
            Assert.Equal(2, secondRoundJob?.VisualRevisionRound);
            Assert.Equal("visual_refinement_limit_reached", thirdRound.Code);
            Assert.Equal("visual_deck_job_superseded", staleBranch.Code);
            Assert.Equal("visual_refinement_must_be_single_slide", multiSlide.Code);
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
    public async Task ConcurrentVisualMutationsDoNotBranchFromTheSameSourceJob()
    {
        var storageRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-jobs-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        Directory.CreateDirectory(uploadsRoot);

        try
        {
            var (service, repository) = CreateJobService(storageRoot, uploadsRoot);
            var caller = new CallerContext("user-1", "conversation-1", null);
            var root = await service.SubmitVisualDeckAsync(
                caller,
                new VisualDeckSpec("排他", [new VisualSlideSpec(VisualSlideKind.Title, "表紙")]),
                false,
                CancellationToken.None);
            await SetStateAsync(repository, root.JobId, JobState.Succeeded);

            async Task<string> RefineAsync(string title)
            {
                try
                {
                    await service.SubmitRefineVisualDeckAsync(
                        caller,
                        root.JobId,
                        [new VisualSlideRevision(1, new VisualSlideSpec(VisualSlideKind.Title, title))],
                        CancellationToken.None);
                    return "accepted";
                }
                catch (PptxValidationException exception)
                {
                    return exception.Code;
                }
            }

            var results = await Task.WhenAll(RefineAsync("A"), RefineAsync("B"));

            Assert.Single(results, static result => result == "accepted");
            Assert.Single(results, static result => result == "visual_deck_operation_in_progress");
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

    private static Task<JobRecord> SetStateAsync(
        FileJobRepository repository,
        string jobId,
        JobState state) =>
        repository.UpdateAsync(
            jobId,
            current => current with
            {
                State = state,
                ProgressPercent = state == JobState.Succeeded ? 100 : current.ProgressPercent,
                CompletedAt = DateTimeOffset.UtcNow,
            },
            CancellationToken.None);

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
            new TemplateRegistry(options, guard),
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
