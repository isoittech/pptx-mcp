using Microsoft.Extensions.Options;
using PptxMcp.Artifacts;
using PptxMcp.Configuration;
using PptxMcp.Jobs;
using PptxMcp.Presentation;
using PptxMcp.Security;
using PptxMcp.Storage;
using PptxMcp.Tools;

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
builder.Services.AddSingleton<TemplateRegistry>();
builder.Services.AddSingleton<FileJobRepository>();
builder.Services.AddSingleton<JobChannel>();
builder.Services.AddSingleton<JobCancellationRegistry>();
builder.Services.AddSingleton<JobService>();
builder.Services.AddSingleton<VisualDeckDraftService>();
builder.Services.AddSingleton<IPresentationEngine, OpenXmlPresentationEngine>();
builder.Services.AddSingleton<IVisualPresentationEngine, PptxGenJsVisualPresentationEngine>();
builder.Services.AddSingleton<PresentationAnalysisCache>();
builder.Services.AddSingleton<LibreOfficeRenderer>();
builder.Services.AddHostedService<DefaultTemplateWarmupService>();
builder.Services.AddHostedService<JobWorker>();
builder.Services.AddHostedService<RetentionWorker>();
builder.Services.AddMcpServer(options =>
    options.ServerInstructions = PowerPointServerInstructions.Build(configuredOptions))
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
