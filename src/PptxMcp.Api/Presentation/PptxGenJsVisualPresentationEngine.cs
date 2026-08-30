using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Presentation;

public sealed class PptxGenJsVisualPresentationEngine(
    IOptions<PptxMcpOptions> options,
    ImageAssetRepository imageAssets) : IVisualPresentationEngine
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = false,
    };

    private readonly string rendererPath = options.Value.VisualRendererPath;
    private readonly TimeSpan timeout = TimeSpan.FromMinutes(options.Value.JobTimeoutMinutes);

    public async Task<VisualDeckCreationResult> CreateAsync(
        string destinationPath,
        VisualDeckSpec deck,
        bool useTemplateChrome,
        CancellationToken cancellationToken)
    {
        var workingDirectory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("The presentation output directory is missing.");
        Directory.CreateDirectory(workingDirectory);
        var specificationPath = Path.Combine(workingDirectory, "visual-deck.json");
        var specification = JsonSerializer.SerializeToNode(deck, SerializerOptions)?.AsObject()
            ?? throw new PptxValidationException("invalid_visual_deck", "The visual deck specification could not be serialized.");
        specification["templateChrome"] = useTemplateChrome;
        var assetMetadata = new JsonObject();
        foreach (var assetId in deck.Slides
                     .Where(static slide => slide.Media is not null)
                     .Select(static slide => slide.Media!.AssetId)
                     .Distinct(StringComparer.Ordinal))
        {
            var asset = imageAssets.Get(assetId);
            assetMetadata[assetId] = new JsonObject
            {
                ["width"] = asset.Width,
                ["height"] = asset.Height,
                ["altText"] = asset.AltText,
                ["sha256"] = asset.Sha256,
            };
        }
        specification["imageAssets"] = assetMetadata;
        await using (var stream = File.Open(specificationPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            await JsonSerializer.SerializeAsync<JsonNode>(stream, specification, SerializerOptions, cancellationToken).ConfigureAwait(false);
        }

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "node",
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            },
        };
        process.StartInfo.ArgumentList.Add(rendererPath);
        process.StartInfo.ArgumentList.Add(specificationPath);
        process.StartInfo.ArgumentList.Add(destinationPath);
        process.StartInfo.Environment["PPTX_MCP_IMAGE_ASSET_ROOT"] = imageAssets.AssetsRoot;

        if (!process.Start())
        {
            throw new PptxValidationException("visual_renderer_failed", "Could not start the visual presentation renderer.");
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);
        using var timeoutSource = new CancellationTokenSource(timeout);
        using var linkedSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutSource.Token);
        try
        {
            await process.WaitForExitAsync(linkedSource.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }

            throw;
        }

        var standardOutput = await standardOutputTask.ConfigureAwait(false);
        var standardError = await standardErrorTask.ConfigureAwait(false);
        if (process.ExitCode != 0 || !File.Exists(destinationPath))
        {
            var diagnostic = string.Join(' ', standardError, standardOutput).Trim();
            if (diagnostic.Length > 2_000)
            {
                diagnostic = diagnostic[..2_000];
            }

            throw new PptxValidationException(
                "visual_renderer_failed",
                $"The visual presentation renderer failed with exit code {process.ExitCode}: {diagnostic}");
        }

        PptxGenJsOpenXmlNormalizer.NormalizeAndValidate(destinationPath);

        var rendererContract = deck.RendererContract ?? "visual-v4";
        var usedDomRenderer = standardOutput
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Contains("PPTX_MCP_RENDERER=dom-to-pptx@2.1.1+react-icons@5.7.0", StringComparer.Ordinal);
        var renderer = usedDomRenderer
            ? $"dom-to-pptx 2.1.1 HTML/CSS renderer + react-icons 5.7.0 Lucide allowlist ({rendererContract})"
            : $"PptxGenJS 4.0.1 declarative fallback renderer {rendererContract}";
        return new VisualDeckCreationResult(
            deck.Slides.Count,
            deck.Slides.Select(slide => slide.Kind.ToString()).ToArray(),
            useTemplateChrome ? $"{renderer} + template chrome" : renderer,
            deck.Slides.Count(static slide => slide.SpeakerNotes is not null),
            VisualDeckValidator.GetDesignWarnings(deck));
    }
}
