using System.Diagnostics;
using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Presentation;

public sealed class PptxGenJsVisualPresentationEngine(IOptions<PptxMcpOptions> options) : IVisualPresentationEngine
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
        CancellationToken cancellationToken)
    {
        var workingDirectory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("The presentation output directory is missing.");
        Directory.CreateDirectory(workingDirectory);
        var specificationPath = Path.Combine(workingDirectory, "visual-deck.json");
        await using (var stream = File.Open(specificationPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            await JsonSerializer.SerializeAsync(stream, deck, SerializerOptions, cancellationToken).ConfigureAwait(false);
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

        return new VisualDeckCreationResult(
            deck.Slides.Count,
            deck.Slides.Select(slide => slide.Kind.ToString()).ToArray(),
            "PptxGenJS 4.0.1 declarative renderer");
    }
}
