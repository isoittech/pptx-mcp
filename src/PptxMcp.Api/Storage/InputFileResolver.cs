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
        if (!SafeIdentifier().IsMatch(caller.UserId))
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

        var useLatestUpload = string.Equals(fileId, "latest", StringComparison.OrdinalIgnoreCase);
        if (!useLatestUpload && !SafeIdentifier().IsMatch(fileId))
        {
            throw new PptxValidationException("invalid_file_id", "The source file identifier is invalid.");
        }

        if (!useLatestUpload
            && caller.AttachmentFileIds is not null
            && !caller.AttachmentFileIds.Contains(fileId))
        {
            throw new PptxValidationException("file_not_found", "The uploaded PowerPoint file was not found in the current message attachments.");
        }

        var candidates = useLatestUpload
            ? FindLatestUpload(userDirectory, caller.AttachmentFileIds)
            : Directory
                .EnumerateFiles(userDirectory, $"{fileId}__*", SearchOption.TopDirectoryOnly)
                .Where(IsRegularPptxUpload)
                .Take(2)
                .ToArray();

        if (candidates.Length != 1)
        {
            throw new PptxValidationException(
                candidates.Length == 0 ? "file_not_found" : "ambiguous_file_id",
                "Exactly one regular PPTX upload must match the file identifier.");
        }

        var resolvedFileId = Path.GetFileName(candidates[0]).Split("__", 2, StringSplitOptions.None)[0];
        var input = await packageGuard.ValidateAsync(candidates[0], cancellationToken).ConfigureAwait(false);
        return input with { FileId = resolvedFileId };
    }

    private static string[] FindLatestUpload(
        string userDirectory,
        IReadOnlySet<string>? attachmentFileIds) => Directory
        // Linux glob matching is case-sensitive. Enumerate the directory and let
        // IsRegularPptxUpload apply the case-insensitive extension check so an
        // uploaded *.PPTX file behaves the same as *.pptx.
        .EnumerateFiles(userDirectory, "*", SearchOption.TopDirectoryOnly)
        .Where(IsRegularPptxUpload)
        .Where(path => attachmentFileIds is null || attachmentFileIds.Contains(FileIdFromName(path)))
        .OrderByDescending(path => new FileInfo(path).LastWriteTimeUtc)
        .ThenByDescending(Path.GetFileName, StringComparer.Ordinal)
        .Take(1)
        .ToArray();

    private static string FileIdFromName(string path) =>
        Path.GetFileName(path).Split("__", 2, StringSplitOptions.None)[0];

    private static bool IsRegularPptxUpload(string path)
    {
        if (!string.Equals(Path.GetExtension(path), ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var file = new FileInfo(path);
        var delimiter = file.Name.IndexOf("__", StringComparison.Ordinal);
        return file.Exists && file.LinkTarget is null && delimiter > 0 &&
            SafeIdentifier().IsMatch(file.Name[..delimiter]);
    }

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex SafeIdentifier();
}

public sealed record ValidatedInput(string Path, long Bytes, int SlideCount, string FileId = "");

public sealed class PptxValidationException(string code, string message) : Exception(message)
{
    public string Code { get; } = code;
}
