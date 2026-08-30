using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class VisualDeckDraftServiceTests
{
    private static readonly CallerContext Caller = new("user-a", "conversation-a", "message-a");

    [Fact]
    public void BuildsCompleteDeckFromOrderedBoundedBatches()
    {
        var service = CreateService(maximumSlides: 10);
        var started = service.Begin(
            Caller,
            "段階生成テスト",
            6,
            new VisualThemeSpec("forest"),
            "経営層向け",
            "ja-JP",
            new VisualDesignSpec("executive", "balanced", "geometric"));

        var first = service.AddSlides(
            Caller,
            started.DraftId,
            1,
            CreateSlides(1, 4));
        var second = service.AddSlides(
            Caller,
            started.DraftId,
            5,
            CreateSlides(5, 2));
        var submission = service.AcquireForSubmission(Caller, started.DraftId);

        Assert.Equal(4, first.AcceptedSlideCount);
        Assert.Equal(5, first.NextSlideNumber);
        Assert.Equal(0, second.RemainingSlideCount);
        Assert.NotNull(submission.Deck);
        Assert.Equal("段階生成テスト", submission.Deck.Title);
        Assert.Equal("forest", submission.Deck.Theme?.Preset);
        Assert.Equal("visual-v6-dom", submission.Deck.RendererContract);
        Assert.Equal(Enumerable.Range(1, 6).Select(number => $"Slide {number}"), submission.Deck.Slides.Select(slide => slide.Title));
    }

    [Fact]
    public void RejectsRepeatedOrOutOfOrderBatchWithoutDuplicatingSlides()
    {
        var service = CreateService();
        var started = service.Begin(Caller, "順序検証", 3, null, null, "ja-JP", null);
        service.AddSlides(Caller, started.DraftId, 1, CreateSlides(1, 2));

        var exception = Assert.Throws<PptxValidationException>(() =>
            service.AddSlides(Caller, started.DraftId, 1, CreateSlides(1, 1)));

        Assert.Equal("visual_draft_position_invalid", exception.Code);
    }

    [Fact]
    public void InfersAppendPositionWhenStartSlideNumberIsOmitted()
    {
        var service = CreateService();
        var started = service.Begin(Caller, "自動追番", 3, null, null, "ja-JP", null);

        var first = service.AddSlides(Caller, started.DraftId, null, CreateSlides(1, 2));
        var second = service.AddSlides(Caller, started.DraftId, null, CreateSlides(3, 1));

        Assert.Equal(3, first.NextSlideNumber);
        Assert.Equal(0, second.RemainingSlideCount);
    }

    [Fact]
    public void RepeatedIdenticalStartIsIdempotentAndDifferentStartIsRejected()
    {
        var service = CreateService();
        var started = service.Begin(
            Caller,
            "固定テスト",
            1,
            new VisualThemeSpec("cyber", AccentColor: "#FF00AA"),
            null,
            "ja-JP",
            new VisualDesignSpec("bold", "balanced", "nodes"),
            "latest",
            "auto");

        var repeated = service.Begin(
            Caller,
            "固定テスト",
            1,
            new VisualThemeSpec("cyber", AccentColor: "#FF00AA"),
            null,
            "ja-JP",
            new VisualDesignSpec("bold", "balanced", "nodes"),
            "latest",
            "auto");
        var activeError = Assert.Throws<PptxValidationException>(() =>
            service.Begin(Caller, "別タイトル", 3, null, null, "ja-JP", null));
        service.AddSlides(Caller, started.DraftId, null, CreateSlides(1, 1));
        var mismatch = Assert.Throws<PptxValidationException>(() =>
            service.AcquireForSubmission(Caller, started.DraftId, "default", "auto"));

        Assert.Equal(started.DraftId, repeated.DraftId);
        Assert.Equal("visual_draft_already_active", activeError.Code);
        Assert.Equal("latest", started.TemplateSourceFileId);
        Assert.True(started.CreativeDirectionLocked);
        Assert.Equal("visual_creative_direction_locked", mismatch.Code);
    }

    [Fact]
    public void RejectsFinishUntilExpectedSlideCountIsComplete()
    {
        var service = CreateService();
        var started = service.Begin(Caller, "未完成検証", 3, null, null, "ja-JP", null);
        service.AddSlides(Caller, started.DraftId, 1, CreateSlides(1, 2));

        var exception = Assert.Throws<PptxValidationException>(() =>
            service.AcquireForSubmission(Caller, started.DraftId));

        Assert.Equal("visual_draft_incomplete", exception.Code);
    }

    [Fact]
    public void EnforcesConversationOwnershipAndReturnsExistingJobAfterSubmission()
    {
        var service = CreateService();
        var started = service.Begin(Caller, "所有権検証", 1, null, null, "ja-JP", null);
        service.AddSlides(Caller, started.DraftId, 1, CreateSlides(1, 1));

        var otherCaller = new CallerContext("user-a", "conversation-b", "message-b");
        var ownershipError = Assert.Throws<PptxValidationException>(() =>
            service.AcquireForSubmission(otherCaller, started.DraftId));
        Assert.Equal("visual_draft_not_found", ownershipError.Code);

        var acquired = service.AcquireForSubmission(Caller, started.DraftId);
        Assert.NotNull(acquired.Deck);
        service.MarkSubmitted(Caller, started.DraftId, "0123456789abcdef0123456789abcdef");
        var repeated = service.AcquireForSubmission(Caller, started.DraftId);
        Assert.Null(repeated.Deck);
        Assert.Equal("0123456789abcdef0123456789abcdef", repeated.ExistingJobId);
    }

    [Fact]
    public void RejectsBatchesLargerThanFourSlides()
    {
        var service = CreateService(maximumSlides: 10);
        var started = service.Begin(Caller, "バッチ上限", 5, null, null, "ja-JP", null);

        var exception = Assert.Throws<PptxValidationException>(() =>
            service.AddSlides(Caller, started.DraftId, 1, CreateSlides(1, 5)));

        Assert.Equal("visual_draft_batch_invalid", exception.Code);
    }

    private static VisualDeckDraftService CreateService(int maximumSlides = 50) =>
        new(
            Options.Create(new PptxMcpOptions { MaxSlides = maximumSlides }),
            TimeProvider.System);

    private static VisualSlideSpec[] CreateSlides(int start, int count) =>
        Enumerable.Range(start, count)
            .Select(number => new VisualSlideSpec(VisualSlideKind.Title, $"Slide {number}"))
            .ToArray();
}
