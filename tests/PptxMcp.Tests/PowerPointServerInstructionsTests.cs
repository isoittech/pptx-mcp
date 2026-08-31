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
        Assert.Contains("append 2 to 4 consecutive complete slides", instructions, StringComparison.Ordinal);
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
        Assert.Contains("visible content appears at least 14pt", instructions, StringComparison.Ordinal);
        Assert.Contains("12pt exception", instructions, StringComparison.Ordinal);
        Assert.Contains("enforces at most three rounds", instructions, StringComparison.Ordinal);
        Assert.Contains("speakerNotes with both purpose and talkScript", instructions, StringComparison.Ordinal);
        Assert.Contains("included in the downloaded PPTX", instructions, StringComparison.Ordinal);
        Assert.Contains("hidden chain-of-thought", instructions, StringComparison.Ordinal);
        Assert.Contains("Omit speakerNotes to inherit", instructions, StringComparison.Ordinal);
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
        Assert.Contains("page-level composition", instructions, StringComparison.Ordinal);
        Assert.Contains("CoverageMap", instructions, StringComparison.Ordinal);
        Assert.Contains("TransformationEvidence", instructions, StringComparison.Ordinal);
        Assert.Contains("ArtifactShowcase", instructions, StringComparison.Ordinal);
        Assert.Contains("GanttSchedule", instructions, StringComparison.Ordinal);
        Assert.Contains("do not request shadows", instructions, StringComparison.Ordinal);
        Assert.Contains("react-icons/lu", instructions, StringComparison.Ordinal);
        Assert.Contains("never send React code", instructions, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("preserves its master, logo, footer, and page numbering", instructions, StringComparison.Ordinal);
        Assert.Contains("LibreOffice PPTX-to-PDF", instructions, StringComparison.Ordinal);
        Assert.Contains("successfully producing a PDF is not visual reflection", instructions, StringComparison.Ordinal);
        Assert.Contains("not a pixel-perfect guarantee for Microsoft PowerPoint", instructions, StringComparison.Ordinal);
        Assert.Contains("PowerPoint as the final rendering authority", instructions, StringComparison.Ordinal);
        Assert.Contains("touching or nearly touching independent objects as a failure", instructions, StringComparison.Ordinal);
        Assert.Contains("Compare title hierarchy across all body pages", instructions, StringComparison.Ordinal);
    }

    [Fact]
    public void DomOnlyCompanyTemplateTrialPublishesRendererAndHeadingContract()
    {
        var instructions = PowerPointServerInstructions.Build(new PptxMcpOptions
        {
            DefaultTemplateId = "organization-default",
            DefaultTemplateCoverSampleSlideNumber = 2,
            DefaultTemplateBodySampleSlideNumber = 4,
            DefaultTemplateBodyUsesAccent2Headings = true,
            RequireDomOnlyRenderer = true,
        });

        Assert.Contains("DOM-only visual-renderer trial", instructions, StringComparison.Ordinal);
        Assert.Contains("server rejects a deck that would use the PptxGenJS compatibility renderer", instructions, StringComparison.Ordinal);
        Assert.Contains("NativeDiagram tree and flow are DOM-supported", instructions, StringComparison.Ordinal);
        Assert.Contains("fallback_rendered_slide_count", instructions, StringComparison.Ordinal);
        Assert.Contains("title and its supporting claim in separate fields", instructions, StringComparison.Ordinal);
        Assert.Contains("title at 30pt", instructions, StringComparison.Ordinal);
        Assert.Contains("subtitle as a 16pt native bullet", instructions, StringComparison.Ordinal);
        Assert.Contains("Accent 2 dark blue", instructions, StringComparison.Ordinal);
        Assert.Contains("does not apply to an uploaded alternate template", instructions, StringComparison.Ordinal);
    }

    [Fact]
    public void ModelAuthoredHtmlDeploymentMakesOpusTheSlideDesigner()
    {
        var instructions = PowerPointServerInstructions.Build(new PptxMcpOptions
        {
            DefaultTemplateId = "organization-default",
            DefaultTemplateCoverSampleSlideNumber = 2,
            DefaultTemplateBodySampleSlideNumber = 4,
            DefaultTemplateBodyUsesAccent2Headings = true,
            UseModelAuthoredHtmlRenderer = true,
            RequireDomOnlyRenderer = true,
        });

        Assert.Contains("visual-v7-author-html", instructions, StringComparison.Ordinal);
        Assert.Contains("You, the conversation model, design every slide as a static 1600x900 web page", instructions, StringComparison.Ordinal);
        Assert.Contains("authoredHtml.html", instructions, StringComparison.Ordinal);
        Assert.Contains("authoredHtml.css", instructions, StringComparison.Ordinal);
        Assert.Contains("must not replace your layout with a fixed card, tree, or diagram template", instructions, StringComparison.Ordinal);
        Assert.Contains("CSS grid/flex/absolute positioning", instructions, StringComparison.Ordinal);
        Assert.Contains("data-pptx-icon", instructions, StringComparison.Ordinal);
        Assert.Contains("data-pptx-asset", instructions, StringComparison.Ordinal);
        Assert.Contains("Every CSS selector must start with .slide", instructions, StringComparison.Ordinal);
        Assert.Contains("do not call pptx_prepare_visual_objects", instructions, StringComparison.Ordinal);
        Assert.Contains("do not set slide.visualObjects", instructions, StringComparison.Ordinal);
        Assert.Contains("leave every assetPlan.visual_object_asset_ids list empty", instructions, StringComparison.Ordinal);
        Assert.Contains("Build every tree, flow, cycle, network, timeline, arrow, frame, and callout as real HTML elements", instructions, StringComparison.Ordinal);
        Assert.Contains("roughly 15% spare height", instructions, StringComparison.Ordinal);
        Assert.Contains("replace the complete authoredHtml composition", instructions, StringComparison.Ordinal);
        Assert.Contains("fallback_rendered_slide_count must be zero", instructions, StringComparison.Ordinal);
        Assert.Contains("send 2 to 4 consecutive complete slides", instructions, StringComparison.Ordinal);
        Assert.Contains("Use two slides as the normal batch size", instructions, StringComparison.Ordinal);
        Assert.Contains("Never shorten or omit a page merely to increase the batch size", instructions, StringComparison.Ordinal);
        Assert.Contains("server rejects a smaller batch", instructions, StringComparison.Ordinal);
        Assert.Contains("every authored font-size must be a literal px value of at least 24px", instructions, StringComparison.Ordinal);
        Assert.Contains("data-pptx-role=\"source-meta\"", instructions, StringComparison.Ordinal);
        Assert.Contains("data-pptx-role=\"body-title\"", instructions, StringComparison.Ordinal);
        Assert.Contains("data-pptx-role=\"body-claim\"", instructions, StringComparison.Ordinal);
        Assert.Contains("data-pptx-role=\"cover-title\"", instructions, StringComparison.Ordinal);
        Assert.Contains("data-pptx-role=\"cover-subtitle\"", instructions, StringComparison.Ordinal);
        Assert.Contains("cover must show only slide.title and one concise slide.subtitle", instructions, StringComparison.Ordinal);
        Assert.Contains("Do not add a white panel", instructions, StringComparison.Ordinal);
        Assert.Contains("50px (30pt)", instructions, StringComparison.Ordinal);
        Assert.Contains("27px (about 16pt)", instructions, StringComparison.Ordinal);
        Assert.Contains("body-title must use exactly 50px", instructions, StringComparison.Ordinal);
        Assert.Contains("body-claim plus its li must use exactly 27px", instructions, StringComparison.Ordinal);
        Assert.Contains("Do not add left padding to the required body-claim ul", instructions, StringComparison.Ordinal);
        Assert.Contains("use an 8px vertical gap between them", instructions, StringComparison.Ordinal);
        Assert.Contains("at least 20px of visible whitespace between the claim and the next independent block", instructions, StringComparison.Ordinal);
        Assert.Contains("Reserve a clear lower safe area", instructions, StringComparison.Ordinal);
        Assert.Contains("shorten visible copy and move supporting detail to speaker notes", instructions, StringComparison.Ordinal);
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
