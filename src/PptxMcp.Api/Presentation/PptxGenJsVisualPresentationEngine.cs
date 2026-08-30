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
    private static readonly HashSet<VisualSlideKind> DomSupportedSlideKinds =
    [
        VisualSlideKind.Title,
        VisualSlideKind.Agenda,
        VisualSlideKind.Section,
        VisualSlideKind.Bullets,
        VisualSlideKind.Metrics,
        VisualSlideKind.Comparison,
        VisualSlideKind.Process,
        VisualSlideKind.Timeline,
        VisualSlideKind.Statement,
        VisualSlideKind.Cards,
        VisualSlideKind.Matrix,
        VisualSlideKind.Funnel,
        VisualSlideKind.Roadmap,
        VisualSlideKind.Quote,
        VisualSlideKind.Closing,
        VisualSlideKind.StructuredBrief,
        VisualSlideKind.Scorecard,
        VisualSlideKind.DataTable,
        VisualSlideKind.Media,
        VisualSlideKind.CoverageMap,
        VisualSlideKind.TransformationEvidence,
        VisualSlideKind.ArtifactShowcase,
        VisualSlideKind.GanttSchedule,
    ];

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
        bool useDefaultTemplateCoverOverlay,
        CancellationToken cancellationToken)
    {
        var workingDirectory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("The presentation output directory is missing.");
        Directory.CreateDirectory(workingDirectory);
        var rendererContract = deck.RendererContract ?? "visual-v4";
        var usePageLevelComposition = rendererContract.Equals("visual-v6-dom", StringComparison.OrdinalIgnoreCase)
            && deck.Slides.Any(static slide => DomSupportedSlideKinds.Contains(slide.Kind))
            && deck.Slides.Any(static slide => !DomSupportedSlideKinds.Contains(slide.Kind));
        var usedDomRenderer = false;
        if (usePageLevelComposition)
        {
            var segmentDirectory = Path.Combine(workingDirectory, "visual-render-segments");
            Directory.CreateDirectory(segmentDirectory);
            var segmentPaths = new List<string>(deck.Slides.Count);
            for (var index = 0; index < deck.Slides.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var slideDirectory = Path.Combine(segmentDirectory, $"slide-{index + 1:D2}");
                Directory.CreateDirectory(slideDirectory);
                var segmentPath = Path.Combine(slideDirectory, "slide.pptx");
                var segmentDeck = deck with { Slides = [deck.Slides[index]] };
                var segmentUsedDom = await RenderDeckAsync(
                    segmentPath,
                    segmentDeck,
                    useTemplateChrome,
                    useDefaultTemplateCoverOverlay,
                    index,
                    deck.Slides.Count,
                    cancellationToken).ConfigureAwait(false);
                PptxGenJsOpenXmlNormalizer.NormalizeAndValidate(segmentPath);
                usedDomRenderer |= segmentUsedDom;
                segmentPaths.Add(segmentPath);
            }

            await OpenXmlRenderedDeckComposer.ComposeAsync(
                segmentPaths,
                destinationPath,
                cancellationToken).ConfigureAwait(false);
        }
        else
        {
            usedDomRenderer = await RenderDeckAsync(
                destinationPath,
                deck,
                useTemplateChrome,
                useDefaultTemplateCoverOverlay,
                0,
                deck.Slides.Count,
                cancellationToken).ConfigureAwait(false);
        }

        PptxGenJsOpenXmlNormalizer.NormalizeAndValidate(destinationPath);

        var renderer = usePageLevelComposition
            ? $"page-level DOM/native composition: dom-to-pptx 2.1.1 + react-icons 5.7.0 with PptxGenJS 4.0.1 native fallback ({rendererContract})"
            : usedDomRenderer
                ? $"dom-to-pptx 2.1.1 HTML/CSS renderer + react-icons 5.7.0 Lucide allowlist ({rendererContract})"
                : $"PptxGenJS 4.0.1 declarative fallback renderer {rendererContract}";
        return new VisualDeckCreationResult(
            deck.Slides.Count,
            deck.Slides.Select(slide => slide.Kind.ToString()).ToArray(),
            useTemplateChrome ? $"{renderer} + template chrome" : renderer,
            deck.Slides.Count(static slide => slide.SpeakerNotes is not null),
            VisualDeckValidator.GetDesignWarnings(deck));
    }

    private async Task<bool> RenderDeckAsync(
        string outputPath,
        VisualDeckSpec deck,
        bool useTemplateChrome,
        bool useDefaultTemplateCoverOverlay,
        int slideNumberOffset,
        int deckTotalSlides,
        CancellationToken cancellationToken)
    {
        var workingDirectory = Path.GetDirectoryName(outputPath)
            ?? throw new InvalidOperationException("The presentation output directory is missing.");
        var specificationPath = Path.Combine(workingDirectory, "visual-deck.json");
        var specification = JsonSerializer.SerializeToNode(deck, SerializerOptions)?.AsObject()
            ?? throw new PptxValidationException("invalid_visual_deck", "The visual deck specification could not be serialized.");
        specification["templateChrome"] = useTemplateChrome;
        specification["defaultTemplateCoverOverlay"] = useDefaultTemplateCoverOverlay;
        specification["slideNumberOffset"] = slideNumberOffset;
        specification["deckTotalSlides"] = deckTotalSlides;
        var assetMetadata = new JsonObject();
        var referencedImageAssetIds = deck.Slides
            .SelectMany(static slide => EnumerateImageAssetIds(slide))
            .Distinct(StringComparer.Ordinal);
        foreach (var assetId in referencedImageAssetIds)
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
        process.StartInfo.ArgumentList.Add(outputPath);
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
        if (process.ExitCode != 0 || !File.Exists(outputPath))
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

        return standardOutput
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Contains("PPTX_MCP_RENDERER=dom-to-pptx@2.1.1+react-icons@5.7.0", StringComparer.Ordinal);
    }

    private static IEnumerable<string> EnumerateImageAssetIds(VisualSlideSpec slide)
    {
        if (slide.Media is not null)
        {
            yield return slide.Media.AssetId;
        }

        if (slide.ArtifactShowcase?.Groups is null)
        {
            yield break;
        }

        foreach (var artifact in slide.ArtifactShowcase.Groups.SelectMany(static group => group.Artifacts))
        {
            yield return artifact.AssetId;
        }
    }
}
