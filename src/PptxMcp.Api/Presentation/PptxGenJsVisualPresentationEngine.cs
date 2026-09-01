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
        VisualSlideKind.NativeDiagram,
    ];

    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = false,
    };

    private readonly string rendererPath = options.Value.VisualRendererPath;
    private readonly TimeSpan timeout = TimeSpan.FromMinutes(options.Value.JobTimeoutMinutes);
    private readonly bool requireDomOnlyRenderer = options.Value.RequireDomOnlyRenderer;

    public async Task<VisualDeckCreationResult> CreateAsync(
        string destinationPath,
        VisualDeckSpec deck,
        bool useTemplateChrome,
        bool useDefaultTemplateCoverOverlay,
        bool useDefaultTemplateBodyStyle,
        CancellationToken cancellationToken)
    {
        var workingDirectory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("The presentation output directory is missing.");
        Directory.CreateDirectory(workingDirectory);
        var rendererContract = deck.RendererContract ?? "visual-v4";
        var isServerOwnedDomContract = rendererContract.Equals("visual-v6-dom", StringComparison.OrdinalIgnoreCase);
        var isModelAuthoredHtmlContract = rendererContract.Equals("visual-v7-author-html", StringComparison.OrdinalIgnoreCase);
        var isDomContract = isServerOwnedDomContract || isModelAuthoredHtmlContract;
        var domSupport = deck.Slides
            .Select(slide => CanRenderWithDom(slide, isModelAuthoredHtmlContract))
            .ToArray();
        if (requireDomOnlyRenderer && isDomContract && domSupport.Any(static supported => !supported))
        {
            var unsupportedKinds = deck.Slides
                .Where((_, index) => !domSupport[index])
                .Select(static slide => slide.Kind.ToString())
                .Distinct(StringComparer.Ordinal)
                .OrderBy(static kind => kind, StringComparer.Ordinal);
            throw new PptxValidationException(
                "dom_renderer_required",
                $"This deployment requires dom-to-pptx for every page in renderer contract {rendererContract}. Unsupported kinds: {string.Join(", ", unsupportedKinds)}.");
        }
        var usePageLevelComposition = isDomContract
            && (isModelAuthoredHtmlContract && deck.Slides.Count > 1
                || domSupport.Any(static supported => supported)
                && domSupport.Any(static supported => !supported));
        var usedDomRenderer = false;
        var rendererUsageBySlide = new List<string>(deck.Slides.Count);
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
                bool segmentUsedDom;
                try
                {
                    segmentUsedDom = await RenderDeckAsync(
                        segmentPath,
                        segmentDeck,
                        useTemplateChrome,
                        useDefaultTemplateCoverOverlay,
                        useDefaultTemplateBodyStyle,
                        index,
                        deck.Slides.Count,
                        cancellationToken).ConfigureAwait(false);
                }
                catch (PptxValidationException exception) when (exception.Code == "visual_authored_html_invalid")
                {
                    throw new PptxValidationException(
                        exception.Code,
                        $"Slide {index + 1} model-authored HTML/CSS is invalid: {exception.Message}");
                }
                PptxGenJsOpenXmlNormalizer.NormalizeAndValidate(segmentPath);
                usedDomRenderer |= segmentUsedDom;
                rendererUsageBySlide.Add(segmentUsedDom ? "dom-to-pptx" : "pptxgenjs-fallback");
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
                useDefaultTemplateBodyStyle,
                0,
                deck.Slides.Count,
                cancellationToken).ConfigureAwait(false);
            if (requireDomOnlyRenderer && isDomContract && !usedDomRenderer)
            {
                throw new PptxValidationException(
                    "dom_renderer_required",
                    $"The renderer did not confirm dom-to-pptx for renderer contract {rendererContract}.");
            }
            rendererUsageBySlide.AddRange(deck.Slides.Select(_ => usedDomRenderer ? "dom-to-pptx" : "pptxgenjs-fallback"));
        }

        PptxGenJsOpenXmlNormalizer.NormalizeAndValidate(destinationPath);

        var renderer = usePageLevelComposition
            ? isModelAuthoredHtmlContract
                ? $"page-level model-authored HTML/CSS: dom-to-pptx 2.1.1 + react-icons 5.7.0 ({rendererContract})"
                : $"page-level DOM/native composition: dom-to-pptx 2.1.1 + react-icons 5.7.0 with PptxGenJS 4.0.1 native fallback ({rendererContract})"
            : usedDomRenderer
                ? $"dom-to-pptx 2.1.1 HTML/CSS renderer + react-icons 5.7.0 Lucide allowlist ({rendererContract})"
                : $"PptxGenJS 4.0.1 declarative fallback renderer {rendererContract}";
        return new VisualDeckCreationResult(
            deck.Slides.Count,
            deck.Slides.Select(slide => slide.Kind.ToString()).ToArray(),
            useTemplateChrome ? $"{renderer} + template chrome" : renderer,
            deck.Slides.Count(static slide => slide.SpeakerNotes is not null),
            VisualDeckValidator.GetDesignWarnings(deck),
            rendererUsageBySlide.Count(static usage => usage == "dom-to-pptx"),
            rendererUsageBySlide.Count(static usage => usage == "pptxgenjs-fallback"),
            rendererUsageBySlide);
    }

    private async Task<bool> RenderDeckAsync(
        string outputPath,
        VisualDeckSpec deck,
        bool useTemplateChrome,
        bool useDefaultTemplateCoverOverlay,
        bool useDefaultTemplateBodyStyle,
        int slideNumberOffset,
        int deckTotalSlides,
        CancellationToken cancellationToken)
    {
        var workingDirectory = Path.GetDirectoryName(outputPath)
            ?? throw new InvalidOperationException("The presentation output directory is missing.");
        var isModelAuthoredHtmlContract = string.Equals(
            deck.RendererContract,
            "visual-v7-author-html",
            StringComparison.OrdinalIgnoreCase);
        var specificationPath = Path.Combine(workingDirectory, "visual-deck.json");
        var specification = JsonSerializer.SerializeToNode(deck, SerializerOptions)?.AsObject()
            ?? throw new PptxValidationException("invalid_visual_deck", "The visual deck specification could not be serialized.");
        specification["templateChrome"] = useTemplateChrome;
        specification["defaultTemplateCoverOverlay"] = useDefaultTemplateCoverOverlay;
        specification["defaultTemplateBodyStyle"] = useDefaultTemplateBodyStyle;
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

        for (var attempt = 1; attempt <= 2; attempt++)
        {
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
            if (process.ExitCode == 0 && File.Exists(outputPath))
            {
                return standardOutput
                    .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Contains("PPTX_MCP_RENDERER=dom-to-pptx@2.1.1+react-icons@5.7.0", StringComparer.Ordinal);
            }

            var diagnostic = string.Join(' ', standardError, standardOutput).Trim();
            if (diagnostic.Length > 2_000)
            {
                diagnostic = diagnostic[..2_000];
            }

            if (attempt == 1 && IsTransientBrowserLaunchFailure(diagnostic))
            {
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                await Task.Delay(TimeSpan.FromMilliseconds(250), cancellationToken).ConfigureAwait(false);
                continue;
            }

            var errorCode = isModelAuthoredHtmlContract && IsAuthoredHtmlValidationFailure(diagnostic)
                ? "visual_authored_html_invalid"
                : "visual_renderer_failed";
            throw new PptxValidationException(
                errorCode,
                $"The visual presentation renderer failed with exit code {process.ExitCode}: {diagnostic}");
        }

        throw new UnreachableException();
    }

    private static bool IsAuthoredHtmlValidationFailure(string diagnostic) =>
        diagnostic.Contains("Model-authored", StringComparison.OrdinalIgnoreCase)
        || diagnostic.Contains("default-template cover", StringComparison.OrdinalIgnoreCase)
        || diagnostic.Contains("default-template body", StringComparison.OrdinalIgnoreCase)
        || diagnostic.Contains("data-pptx-role", StringComparison.OrdinalIgnoreCase);

    internal static bool IsTransientBrowserLaunchFailure(string diagnostic) =>
        diagnostic.Contains("Failed to launch headless browser", StringComparison.OrdinalIgnoreCase)
        && diagnostic.Contains("Timed out", StringComparison.OrdinalIgnoreCase)
        && diagnostic.Contains("WS endpoint URL", StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<string> EnumerateImageAssetIds(VisualSlideSpec slide)
    {
        if (slide.AuthoredHtml?.AssetIds is not null)
        {
            foreach (var assetId in slide.AuthoredHtml.AssetIds)
            {
                yield return assetId;
            }
        }

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

    private static bool CanRenderWithDom(VisualSlideSpec slide, bool modelAuthoredHtml)
    {
        if (modelAuthoredHtml)
        {
            return slide.AuthoredHtml is not null;
        }

        if (!DomSupportedSlideKinds.Contains(slide.Kind))
        {
            return false;
        }

        return slide.Kind != VisualSlideKind.NativeDiagram
            || slide.Diagram?.Kind is VisualDiagramKind.Tree or VisualDiagramKind.Flow;
    }
}
