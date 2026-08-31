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

    [Fact]
    public void UsesModelAuthoredHtmlContractOnlyWhenDeploymentEnablesIt()
    {
        var service = CreateService(useModelAuthoredHtmlRenderer: true);
        var started = service.Begin(Caller, "HTML生成", 1, null, null, "ja-JP", null);
        service.AddSlides(
            Caller,
            started.DraftId,
            null,
            [
                new VisualSlideSpec(
                    VisualSlideKind.NativeDiagram,
                    "判断構造",
                    AuthoredHtml: new VisualAuthoredHtmlSpec(
                        "<div class=\"layout\"><h1>判断構造</h1></div>",
                        ".slide{background:white}.slide .layout{display:grid}")),
            ]);

        var submission = service.AcquireForSubmission(Caller, started.DraftId);

        Assert.Equal("visual-v7-author-html", submission.Deck?.RendererContract);
        Assert.NotNull(submission.Deck?.Slides[0].AuthoredHtml);
    }

    [Fact]
    public void ModelAuthoredHtmlContractAcceptsBoundedBatches()
    {
        var service = CreateService(useModelAuthoredHtmlRenderer: true);
        var started = service.Begin(Caller, "HTML小分け追加", 5, null, null, "ja-JP", null);

        Assert.Equal(4, started.MaximumBatchSlides);
        Assert.Contains("next 2 consecutive complete slides", started.Instruction, StringComparison.Ordinal);
        Assert.Contains("Only send 3 to 4", started.Instruction, StringComparison.Ordinal);

        var undersizedException = Assert.Throws<PptxValidationException>(() => service.AddSlides(
            Caller,
            started.DraftId,
            null,
            [CreateAuthoredSlide("1枚目")]));

        Assert.Equal("visual_authored_html_batch_too_small", undersizedException.Code);

        var accepted = service.AddSlides(
            Caller,
            started.DraftId,
            null,
            [
                CreateAuthoredSlide("1枚目"),
                CreateAuthoredSlide("2枚目"),
                CreateAuthoredSlide("3枚目"),
            ]);

        Assert.Equal(3, accepted.AcceptedSlideCount);
        Assert.Contains("next 2 consecutive complete slides", accepted.Instruction, StringComparison.Ordinal);

        var finished = service.AddSlides(
            Caller,
            started.DraftId,
            null,
            [CreateAuthoredSlide("4枚目"), CreateAuthoredSlide("5枚目")]);
        Assert.Equal(0, finished.RemainingSlideCount);

        var oversizedBatchService = CreateService(useModelAuthoredHtmlRenderer: true);
        var oversizedBatchDraft = oversizedBatchService.Begin(Caller, "HTML過大batch", 5, null, null, "ja-JP", null);
        var exception = Assert.Throws<PptxValidationException>(() => oversizedBatchService.AddSlides(
            Caller,
            oversizedBatchDraft.DraftId,
            null,
            [
                CreateAuthoredSlide("1枚目"),
                CreateAuthoredSlide("2枚目"),
                CreateAuthoredSlide("3枚目"),
                CreateAuthoredSlide("4枚目"),
                CreateAuthoredSlide("5枚目"),
            ]));

        Assert.Equal("visual_draft_batch_invalid", exception.Code);
    }

    [Fact]
    public void DefaultTemplateBodyRequiresSeparateSemanticAndHtmlTitleClaim()
    {
        var service = CreateService(
            useModelAuthoredHtmlRenderer: true,
            defaultTemplateId: "organization-default",
            defaultTemplateBodyUsesAccent2Headings: true);
        var started = service.Begin(Caller, "HTML見出し契約", 2, null, null, "ja-JP", null);

        var exception = Assert.Throws<PptxValidationException>(() => service.AddSlides(
            Caller,
            started.DraftId,
            null,
            [
                CreateAuthoredSlide("表紙"),
                CreateAuthoredSlide("本文") with { Subtitle = "主張" },
            ]));

        Assert.Equal("visual_default_body_heading_contract_required", exception.Code);
        Assert.Contains("retry the same complete batch", exception.Message, StringComparison.Ordinal);
        Assert.Contains("no slide in the rejected batch was accepted", exception.Message, StringComparison.Ordinal);

        var accepted = service.AddSlides(
            Caller,
            started.DraftId,
            null,
            [
                CreateAuthoredSlide("表紙"),
                CreateAuthoredSlide("本文", "主張", includeBodyRoles: true),
            ]);
        Assert.Equal(0, accepted.RemainingSlideCount);
    }

    [Fact]
    public void AlternateTemplateDoesNotInheritDefaultBodyHeadingContract()
    {
        var service = CreateService(
            useModelAuthoredHtmlRenderer: true,
            defaultTemplateId: "organization-default",
            defaultTemplateBodyUsesAccent2Headings: true);
        var started = service.Begin(
            Caller,
            "テンプレートなしHTML",
            2,
            null,
            null,
            "ja-JP",
            null,
            "none");

        var accepted = service.AddSlides(
            Caller,
            started.DraftId,
            null,
            [CreateAuthoredSlide("表紙"), CreateAuthoredSlide("本文")]);

        Assert.Equal(0, accepted.RemainingSlideCount);
    }

    private static VisualDeckDraftService CreateService(
        int maximumSlides = 50,
        bool useModelAuthoredHtmlRenderer = false,
        string defaultTemplateId = "",
        bool defaultTemplateBodyUsesAccent2Headings = false) =>
        new(
            Options.Create(new PptxMcpOptions
            {
                MaxSlides = maximumSlides,
                UseModelAuthoredHtmlRenderer = useModelAuthoredHtmlRenderer,
                DefaultTemplateId = defaultTemplateId,
                DefaultTemplateBodyUsesAccent2Headings = defaultTemplateBodyUsesAccent2Headings,
            }),
            TimeProvider.System);

    private static VisualSlideSpec CreateAuthoredSlide(
        string title,
        string? subtitle = null,
        bool includeBodyRoles = false) =>
        new(
            VisualSlideKind.StructuredBrief,
            title,
            Subtitle: subtitle,
            AuthoredHtml: new VisualAuthoredHtmlSpec(
                includeBodyRoles
                    ? $"<div><h2 data-pptx-role=\"body-title\">{title}</h2><ul data-pptx-role=\"body-claim\"><li>{subtitle}</li></ul></div>"
                    : $"<div><h2>{title}</h2></div>",
                ".slide{padding:64px}.slide h2{font-size:40px}"));

    private static VisualSlideSpec[] CreateSlides(int start, int count) =>
        Enumerable.Range(start, count)
            .Select(number => new VisualSlideSpec(VisualSlideKind.Title, $"Slide {number}"))
            .ToArray();
}
