using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;

namespace PptxMcp.Storage;

public sealed partial class InputFileResolver(
    IOptions<PptxMcpOptions> options,
    PptxPackageGuard packageGuard)
{
    private readonly PptxMcpOptions options = options.Value;

    public async Task<ValidatedInput> ResolveAsync(
        CallerContext caller,
        string fileId,
        CancellationToken cancellationToken)
    {
        if (!SafeIdentifier().IsMatch(fileId) || !SafeIdentifier().IsMatch(caller.UserId))
        {
            throw new PptxValidationException("invalid_file_id", "The source file identifier is invalid.");
        }

        var userDirectory = Path.GetFullPath(Path.Combine(options.LibreChatUploadsRoot, caller.UserId));
        var uploadsRoot = Path.GetFullPath(options.LibreChatUploadsRoot) + Path.DirectorySeparatorChar;
        if (!userDirectory.StartsWith(uploadsRoot, StringComparison.Ordinal))
        {
            throw new PptxValidationException("invalid_file_id", "The source file identifier is invalid.");
        }

        if (!Directory.Exists(userDirectory))
        {
            throw new PptxValidationException("file_not_found", "The uploaded PowerPoint file was not found.");
        }

        var candidates = Directory
            .EnumerateFiles(userDirectory, $"{fileId}__*", SearchOption.TopDirectoryOnly)
            .Where(path => string.Equals(Path.GetExtension(path), ".pptx", StringComparison.OrdinalIgnoreCase))
            .Take(2)
            .ToArray();

        if (candidates.Length != 1 || new FileInfo(candidates[0]).LinkTarget is not null)
        {
            throw new PptxValidationException(
                candidates.Length == 0 ? "file_not_found" : "ambiguous_file_id",
                "Exactly one regular PPTX upload must match the file identifier.");
        }

        return await packageGuard.ValidateAsync(candidates[0], cancellationToken).ConfigureAwait(false);
    }

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex SafeIdentifier();
}

public sealed record ValidatedInput(string Path, long Bytes, int SlideCount);

public sealed class PptxValidationException(string code, string message) : Exception(message)
{
    public string Code { get; } = code;
}
