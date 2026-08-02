using System.Diagnostics;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Storage;

namespace PptxMcp.Presentation;

public sealed class LibreOfficeRenderer(IOptions<PptxMcpOptions> options)
{
    private readonly TimeSpan timeout = TimeSpan.FromMinutes(options.Value.JobTimeoutMinutes);

    public async Task<IReadOnlyList<string>> RenderAsync(
        string sourcePath,
        string outputDirectory,
        CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(outputDirectory);
        var profileDirectory = Path.Combine(outputDirectory, ".lo-profile");
        Directory.CreateDirectory(profileDirectory);
        var profileUri = new Uri(profileDirectory + Path.DirectorySeparatorChar).AbsoluteUri;

        await RunAsync(
            "soffice",
            ["--headless", "--nologo", "--nodefault", "--nolockcheck", $"-env:UserInstallation={profileUri}", "--convert-to", "pdf", "--outdir", outputDirectory, sourcePath],
            outputDirectory,
            cancellationToken).ConfigureAwait(false);

        var pdfPath = Directory.EnumerateFiles(outputDirectory, "*.pdf", SearchOption.TopDirectoryOnly).SingleOrDefault()
            ?? throw new PptxValidationException("render_failed", "LibreOffice did not produce a PDF preview.");

        var imagePrefix = Path.Combine(outputDirectory, "slide");
        await RunAsync(
            "pdftoppm",
            ["-png", "-r", "110", pdfPath, imagePrefix],
            outputDirectory,
            cancellationToken).ConfigureAwait(false);

        Directory.Delete(profileDirectory, recursive: true);
        File.Delete(pdfPath);
        var images = Directory.EnumerateFiles(outputDirectory, "slide-*.png", SearchOption.TopDirectoryOnly)
            .Order(StringComparer.Ordinal)
            .ToArray();
        if (images.Length == 0)
        {
            throw new PptxValidationException("render_failed", "No slide preview images were produced.");
        }

        return images;
    }

    private async Task RunAsync(
        string fileName,
        IReadOnlyList<string> arguments,
        string workingDirectory,
        CancellationToken cancellationToken)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            },
        };
        foreach (var argument in arguments)
        {
            process.StartInfo.ArgumentList.Add(argument);
        }

        if (!process.Start())
        {
            throw new PptxValidationException("render_failed", $"Could not start {fileName}.");
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
            process.Kill(entireProcessTree: true);
            throw;
        }

        var standardOutput = await standardOutputTask.ConfigureAwait(false);
        var standardError = await standardErrorTask.ConfigureAwait(false);
        if (process.ExitCode != 0)
        {
            throw new PptxValidationException(
                "render_failed",
                $"{fileName} exited with code {process.ExitCode}: {standardError} {standardOutput}".Trim());
        }
    }
}
