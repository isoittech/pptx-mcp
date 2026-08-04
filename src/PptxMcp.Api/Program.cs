using Microsoft.Extensions.Options;
using PptxMcp.Artifacts;
using PptxMcp.Configuration;
using PptxMcp.Jobs;
using PptxMcp.Presentation;
using PptxMcp.Security;
using PptxMcp.Storage;

var builder = WebApplication.CreateBuilder(args);
var configuredOptions = builder.Configuration
    .GetSection(PptxMcpOptions.SectionName)
    .Get<PptxMcpOptions>() ?? new PptxMcpOptions();
configuredOptions.Validate(requireSecrets: true);

builder.Services.AddSingleton<IOptions<PptxMcpOptions>>(Options.Create(configuredOptions));
builder.Services.AddSingleton(TimeProvider.System);
builder.Services.AddHttpContextAccessor();
builder.Services.AddSingleton<CallerContextAccessor>();
builder.Services.AddSingleton<ArtifactTokenService>();
builder.Services.AddSingleton<RetentionPolicy>();
builder.Services.AddSingleton<PptxPackageGuard>();
builder.Services.AddSingleton<InputFileResolver>();
builder.Services.AddSingleton<FileJobRepository>();
builder.Services.AddSingleton<JobChannel>();
builder.Services.AddSingleton<JobCancellationRegistry>();
builder.Services.AddSingleton<JobService>();
builder.Services.AddSingleton<IPresentationEngine, OpenXmlPresentationEngine>();
builder.Services.AddSingleton<IVisualPresentationEngine, PptxGenJsVisualPresentationEngine>();
builder.Services.AddSingleton<LibreOfficeRenderer>();
builder.Services.AddHostedService<JobWorker>();
builder.Services.AddHostedService<RetentionWorker>();
builder.Services.AddMcpServer(options => options.ServerInstructions = """
    For a new presentation based on an uploaded corporate template, first call pptx_analyze and then pptx_get_job. Build the complete final slide list before calling pptx_create_deck. The slides argument is behaviorally required and must contain every requested slide in one call; never call pptx_create_deck with only sourceFileId. If an empty call returns input_required, immediately retry the same tool with all slides. In each slide use the exact snake_case keys layout_id and fields. In each field use text and the exact shape_id from the analyzed layout, or shape_name/placeholder_index only when needed. Copy layout_id, shape_id, and placeholder_index values verbatim from the analysis result. Never invent, abbreviate, correct, or otherwise modify a layout path or placeholder identifier.

    For new presentations without an uploaded template, use pptx_create_visual_deck and choose the semantic layout that best matches each slide. Keep one message per slide and respect the content limits in the tool schema. After any create or edit job succeeds, you MUST call pptx_get_preview_images for every slide in batches of up to four. Inspect clipping, overflow, overlap, alignment, spacing, contrast, hierarchy, density, balance, and cross-slide consistency. If a template-based deck needs correction, call pptx_refine_deck with the successful jobId and only the changed slides; never resend the complete deck to pptx_create_deck. For a visual deck, revise its declarative specification. Perform at most two correction rounds. Only provide the PPTX download link after this visual review. Never claim visual review was completed unless preview images were actually retrieved.
    """)
    .WithHttpTransport(options => options.Stateless = true)
    .WithToolsFromAssembly();

var app = builder.Build();
app.UseMiddleware<SharedSecretMiddleware>();
app.MapGet("/health/live", () => Results.Ok(new { status = "ok" }));
app.MapGet("/health/ready", (FileJobRepository _) => Results.Ok(new { status = "ready" }));
app.MapArtifactEndpoints();
app.MapMcp("/mcp");
app.Run();

public partial class Program;
