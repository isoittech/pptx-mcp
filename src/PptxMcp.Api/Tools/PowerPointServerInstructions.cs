using PptxMcp.Configuration;

namespace PptxMcp.Tools;

public static class PowerPointServerInstructions
{
    public static string Build(PptxMcpOptions options)
    {
        var defaultTemplateInstructions = string.IsNullOrWhiteSpace(options.DefaultTemplateId)
            ? """
                No deployment default template is configured. For an ordinary new presentation, call pptx_start_visual_deck with templateSourceFileId=none, then complete the staged visual-draft workflow. For a new presentation based on an uploaded corporate template, first call pptx_analyze and pptx_wait_for_job, then call pptx_start_visual_deck with templateSourceFileId set to the uploaded sourceFileId (or latest when unavailable) and templateLayoutId=auto. Template selection is locked at start; omit template arguments from the finish call.
                """
            : """
                A deployment default template is configured. For ordinary new presentations, leave templateSourceFileId=default and templateLayoutId=auto on pptx_start_visual_deck; the MCP server applies the default template without pptx_analyze. Do not ask the user to upload or analyze the default template for ordinary visual generation. If the user explicitly requests strict placement into the default template's existing placeholders, call pptx_analyze with sourceFileId=default and then pptx_wait_for_job to retrieve the startup-cached layout identifiers. If the user explicitly requests no template, set templateSourceFileId=none when starting. If the user explicitly requests an attached alternate template, first call pptx_analyze and pptx_wait_for_job, then set templateSourceFileId to that upload (or latest) and templateLayoutId=auto on pptx_start_visual_deck. Template selection is locked at start and an alternate template overrides the deployment default only for that workflow.
                """;

        var firstNoticeInstructions = string.IsNullOrWhiteSpace(options.FirstAssistantNotice)
            ? string.Empty
            : $"""

                This deployment requires the following one-time user notice:
                <first-assistant-notice>
                {options.FirstAssistantNotice.Trim()}
                </first-assistant-notice>
                Emit this notice only while responding directly to the first user message, before any tool call or tool result exists in that turn. Put it verbatim at the start of the first user-visible assistant text, before any explanation. Never emit the notice in a continuation produced after receiving any tool result, even if earlier assistant text is absent from the continuation context. Never repeat it in later turns, and never show the XML-like delimiter tags.
                """;

        return $$"""
            {{defaultTemplateInstructions}}

            For every new polished visual presentation, first decide the complete outline, exact final slide count, template selection, theme, and design. Call pptx_start_visual_deck exactly once with those choices; they are locked for the workflow. Omit startSlideNumber when calling pptx_add_visual_slides_to_draft so the server appends the next 1-4 complete slides at the current end. Continue until remaining_slide_count is zero, never resend an accepted batch, and never finish an incomplete draft. Finally call exactly one finish tool with the draftId and do not change template or creative direction at finish. Never call these tools with empty arguments. Once a visual deck succeeds, never call pptx_start_visual_deck or a finish tool again merely to improve appearance: use page-level refinement. The server permits only one recovery restart after a failed initial generation. Set userRequestedNewWorkflow=true only when the user explicitly asks for a separate new deck after a prior successful deck. Keep one message per slide, use at least four different visual layouts in decks of six or more slides, and use bullets only when they are the clearest structure. For a text-rich business page, use StructuredBrief with two or three sections and design.density=detailed instead of compressing one long body. For a comparison across evaluation criteria and options, use Scorecard. When the user requests standard notation or ukulele TAB, use MusicScore with matching pitch, string, and fret data. Most sections should remain neutral; reserve strong color and highlight labels for exceptions or decisions.

            For a polished deck with an uploaded alternate template, select that source and templateLayoutId=auto at pptx_start_visual_deck, then finish with pptx_finish_branded_visual_deck without changing those values. It preserves the template master/logo/footer and adds editable visual layouts. Explicit theme colors and fontFace selected at start take priority; template values fill only unspecified theme fields. Use pptx_create_deck only when the user explicitly needs strict placement into existing placeholders. For strict placement, build the complete final slide list before calling pptx_create_deck. The slides argument is behaviorally required and must contain every requested slide in one call; never call pptx_create_deck with only sourceFileId. If an empty call returns input_required, immediately retry the same tool with all slides. Use exact snake_case keys and copy layout_id, shape_id, and placeholder_index values verbatim from analysis.

            When the user asks to add, append, or insert slides into a successful visual or branded visual deck, call pptx_insert_visual_slides with jobId=latest and only the new slides. Omit afterSlideNumber to append, or set it to the existing one-based slide number after which the new slides belong. Never reconstruct or resend the existing slides and never start a new draft for this operation. After an asynchronous tool returns a queued receipt, call pptx_wait_for_job once instead of repeatedly calling pptx_get_job. After any create or edit job succeeds, you MUST call pptx_get_preview_images for every slide in batches of up to four. Inspect clipping, overflow, overlap, alignment, spacing, contrast, hierarchy, density, balance, and cross-slide consistency. Also confirm that headings alone tell the story, reading order is unambiguous, each block carries one main point, strong emphasis occupies roughly 15% or less, and content text does not appear below 9 pt. If a strict placeholder template deck needs correction, call pptx_refine_deck with only changed slides. For a visual or branded visual deck, call pptx_refine_visual_slide with exactly one complete replacement slide and jobId=latest. Revisions accumulate in one job lineage; stale branch jobs are rejected. The server advances review rounds when the same page is refined again and enforces at most two rounds. Never use a multi-page revisions array, restart the complete deck, or raise a recursion limit to continue visual changes. Use pptx_insert_visual_slides when slide count must increase. Only provide the PPTX download link after visual review, and never claim review was completed unless preview images were retrieved.{{firstNoticeInstructions}}
            """;
    }
}
