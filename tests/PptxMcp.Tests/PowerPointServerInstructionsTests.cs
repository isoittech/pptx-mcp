using PptxMcp.Configuration;
using PptxMcp.Tools;

namespace PptxMcp.Tests;

public sealed class PowerPointServerInstructionsTests
{
    [Fact]
    public void ConfiguredDefaultTemplateAvoidsRuntimeAnalysis()
    {
        var instructions = PowerPointServerInstructions.Build(new PptxMcpOptions
        {
            DefaultTemplateId = "organization-default",
        });

        Assert.Contains("without pptx_analyze", instructions, StringComparison.Ordinal);
        Assert.Contains("templateSourceFileId=none", instructions, StringComparison.Ordinal);
        Assert.Contains("Template selection is locked at start", instructions, StringComparison.Ordinal);
        Assert.Contains("call pptx_wait_for_job once", instructions, StringComparison.Ordinal);
        Assert.Contains("instead of repeatedly calling pptx_get_job", instructions, StringComparison.Ordinal);
        Assert.Contains("pptx_insert_visual_slides with jobId=latest and only the new slides", instructions, StringComparison.Ordinal);
        Assert.Contains("Never reconstruct or resend the existing slides", instructions, StringComparison.Ordinal);
        Assert.Contains("never start a new draft merely to bypass the insertion restriction", instructions, StringComparison.Ordinal);
        Assert.Contains("pptx_start_visual_deck exactly once", instructions, StringComparison.Ordinal);
        Assert.Contains("Omit startSlideNumber", instructions, StringComparison.Ordinal);
        Assert.Contains("next 1-4 complete slides", instructions, StringComparison.Ordinal);
        Assert.Contains("remaining_slide_count is zero", instructions, StringComparison.Ordinal);
        Assert.Contains("only one recovery restart", instructions, StringComparison.Ordinal);
        Assert.Contains("openxml_validation_failed", instructions, StringComparison.Ordinal);
        Assert.Contains("do not claim that shortening text", instructions, StringComparison.Ordinal);
        Assert.Contains("userRequestedNewWorkflow=true", instructions, StringComparison.Ordinal);
        Assert.Contains("StructuredBrief", instructions, StringComparison.Ordinal);
        Assert.Contains("design.density=detailed", instructions, StringComparison.Ordinal);
        Assert.Contains("Scorecard", instructions, StringComparison.Ordinal);
        Assert.Contains("MusicScore", instructions, StringComparison.Ordinal);
        Assert.Contains("matching pitch, string, and fret", instructions, StringComparison.Ordinal);
        Assert.Contains("headings alone tell the story", instructions, StringComparison.Ordinal);
        Assert.Contains("roughly 15% or less", instructions, StringComparison.Ordinal);
        Assert.Contains("below 9 pt", instructions, StringComparison.Ordinal);
        Assert.Contains("enforces at most two rounds", instructions, StringComparison.Ordinal);
        Assert.Contains("speakerNotes with both purpose and talkScript", instructions, StringComparison.Ordinal);
        Assert.Contains("included in the downloaded PPTX", instructions, StringComparison.Ordinal);
        Assert.Contains("hidden chain-of-thought", instructions, StringComparison.Ordinal);
        Assert.Contains("omit speakerNotes to inherit", instructions, StringComparison.Ordinal);
        Assert.Contains("For a text-only edit such as translation", instructions, StringComparison.Ordinal);
        Assert.Contains("without any model-controlled override", instructions, StringComparison.Ordinal);
        Assert.Contains("pptx_analyze with includeLayouts=false", instructions, StringComparison.Ordinal);
        Assert.Contains("fixed slides tuples [slide_number, exact_texts[]]", instructions, StringComparison.Ordinal);
        Assert.Contains("table cells are returned as separate exact_texts entries", instructions, StringComparison.Ordinal);
        Assert.Contains("including table body cells", instructions, StringComparison.Ordinal);
        Assert.Contains("When analysis_truncated=false", instructions, StringComparison.Ordinal);
        Assert.Contains("charts field means editable PowerPoint charts", instructions, StringComparison.Ordinal);
        Assert.Contains("If charts=false", instructions, StringComparison.Ordinal);
        Assert.Contains("part of a picture or flattened graphic", instructions, StringComparison.Ordinal);
        Assert.Contains("never describe visible chart-like graphics as chart data", instructions, StringComparison.Ordinal);
        Assert.Contains("do not ask them to choose translation", instructions, StringComparison.Ordinal);
        Assert.Contains("Do not call pptx_get_job after a successful wait", instructions, StringComparison.Ordinal);
        Assert.Contains("at most 20 entries each", instructions, StringComparison.Ordinal);
        Assert.Contains("exact returned job_id as previousJobId", instructions, StringComparison.Ordinal);
        Assert.Contains("Never omit previousJobId again", instructions, StringComparison.Ordinal);
        Assert.Contains("isFinalBatch=true only on the last batch", instructions, StringComparison.Ordinal);
        Assert.Contains("Do not call pptx_render_preview or pptx_get_preview_images", instructions, StringComparison.Ordinal);
        Assert.Contains("retrieve that final job's preview images", instructions, StringComparison.Ordinal);
        Assert.Contains("consecutive groups of four slide numbers", instructions, StringComparison.Ordinal);
        Assert.Contains("never request the same slide twice", instructions, StringComparison.Ordinal);
        Assert.Contains("Do not end the turn after source analysis", instructions, StringComparison.Ordinal);
        Assert.Contains("visual-v6-dom", instructions, StringComparison.Ordinal);
        Assert.Contains("dom-to-pptx", instructions, StringComparison.Ordinal);
        Assert.Contains("react-icons/lu", instructions, StringComparison.Ordinal);
        Assert.Contains("never send React code", instructions, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("preserves its master, logo, footer, and page numbering", instructions, StringComparison.Ordinal);
        Assert.Contains("LibreOffice PPTX-to-PDF", instructions, StringComparison.Ordinal);
        Assert.Contains("successfully producing a PDF is not visual reflection", instructions, StringComparison.Ordinal);
        Assert.Contains("not a pixel-perfect guarantee for Microsoft PowerPoint", instructions, StringComparison.Ordinal);
    }

    [Fact]
    public void DeploymentNoticeIsInjectedVerbatimAndOnlyWhenConfigured()
    {
        const string notice = "PowerPoint資料には既定テンプレートを使用します。";

        var configured = PowerPointServerInstructions.Build(new PptxMcpOptions
        {
            FirstAssistantNotice = notice,
        });
        var unconfigured = PowerPointServerInstructions.Build(new PptxMcpOptions());

        Assert.Contains(notice, configured, StringComparison.Ordinal);
        Assert.Contains("one-time user notice", configured, StringComparison.Ordinal);
        Assert.Contains("Never emit the notice in a continuation produced after receiving any tool result", configured, StringComparison.Ordinal);
        Assert.DoesNotContain("one-time user notice", unconfigured, StringComparison.Ordinal);
    }
}
