using PptxMcp.Configuration;

namespace PptxMcp.Tools;

public static class PowerPointServerInstructions
{
    public static string Build(PptxMcpOptions options)
    {
        var defaultTemplateInstructions = string.IsNullOrWhiteSpace(options.DefaultTemplateId)
            ? """
                No deployment default template is configured. For a new presentation based on an uploaded corporate template, first call pptx_analyze and then pptx_wait_for_job. Use pptx_create_branded_visual_deck for polished visual output, and pass the uploaded sourceFileId (or latest when the identifier is unavailable).
                """
            : """
                A deployment default template is configured. For ordinary new presentations, call pptx_create_visual_deck and leave useDefaultTemplate at its default true value; the MCP server applies the default template without pptx_analyze. Do not ask the user to upload the default template and do not analyze it for ordinary visual generation. If the user explicitly requests strict placement into the default template's existing placeholders, call pptx_analyze with sourceFileId=default and then pptx_wait_for_job to retrieve the startup-cached layout identifiers. If the user explicitly requests no template, set useDefaultTemplate=false. If the user explicitly requests an attached alternate template, first call pptx_analyze and pptx_wait_for_job for that upload, then call pptx_create_branded_visual_deck with its sourceFileId (or latest when the identifier is unavailable). An explicit alternate template overrides the deployment default only for that operation.
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

            If the user wants a polished, visual, infographic, or executive-quality deck with an uploaded alternate template, use pptx_create_branded_visual_deck. It preserves the template master/logo/footer, applies template colors and fonts automatically, and adds editable visual layouts. Use templateLayoutId=auto unless the user explicitly needs a specific zero-placeholder blank layout. Use pptx_create_deck only when the user explicitly needs strict placement into a template's existing placeholders. For strict placement, build the complete final slide list before calling pptx_create_deck. The slides argument is behaviorally required and must contain every requested slide in one call; never call pptx_create_deck with only sourceFileId. If an empty call returns input_required, immediately retry the same tool with all slides. In each slide use the exact snake_case keys layout_id and fields. In each field use text and the exact shape_id from the analyzed layout, or shape_name/placeholder_index only when needed. Copy layout_id, shape_id, and placeholder_index values verbatim from the analysis result. Never invent, abbreviate, correct, or otherwise modify a layout path or placeholder identifier.

            For a new visual presentation, use pptx_create_visual_deck and choose the semantic layout that best matches each slide. Keep one message per slide and respect the content limits in the tool schema. When the user asks to add, append, or insert slides into a successful visual or branded visual deck, call pptx_insert_visual_slides with jobId=latest and only the new slides. Omit afterSlideNumber to append, or set it to the existing one-based slide number after which the new slides belong. Never reconstruct or resend the existing slides and never call a create tool for this operation. The insertion tool preserves the existing deck specification, design, and template server-side and normally waits for completion; if it returns Succeeded, do not call a job status tool. After an asynchronous tool returns a queued receipt, call pptx_wait_for_job once instead of repeatedly calling pptx_get_job. If the wait result is still Queued or Running, call pptx_wait_for_job again with the same jobId. After any create or edit job succeeds, you MUST call pptx_get_preview_images for every slide in batches of up to four. Inspect clipping, overflow, overlap, alignment, spacing, contrast, hierarchy, density, balance, and cross-slide consistency. If a strict placeholder template deck needs correction, call pptx_refine_deck with the successful jobId and only the changed slides; never resend the complete deck to pptx_create_deck. For a visual or branded visual deck, prefer pptx_refine_visual_slide: pass exactly one complete replacement slide per call and use jobId=latest so revisions accumulate without a large tool payload. The refinement tool replaces only an existing slide; use pptx_insert_visual_slides when the slide count must increase. The tool normally waits and returns a terminal status; when it returns Succeeded, do not call pptx_get_job or pptx_wait_for_job and immediately refine the next problem slide. Use pptx_refine_visual_deck only when a client can reliably send the complete revisions array. Perform at most two review rounds. Only provide the PPTX download link after this visual review. Never claim visual review was completed unless preview images were actually retrieved.{{firstNoticeInstructions}}
            """;
    }
}
