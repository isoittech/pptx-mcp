using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;

namespace PptxMcp.Storage;

public sealed partial class UploadedImageResolver(IOptions<PptxMcpOptions> options)
{
    private static readonly byte[] PngSignature = [137, 80, 78, 71, 13, 10, 26, 10];
    private readonly PptxMcpOptions options = options.Value;

    public async Task<ValidatedImageUpload> ResolveAsync(
        CallerContext caller,
        string fileId,
        CancellationToken cancellationToken)
    {
        if (!SafeIdentifier().IsMatch(caller.UserId))
        {
            throw InvalidFileId();
        }

        var uploadsRoot = Path.GetFullPath(options.LibreChatUploadsRoot);
        var rootPrefix = uploadsRoot + Path.DirectorySeparatorChar;
        var userDirectory = Path.GetFullPath(Path.Combine(uploadsRoot, caller.UserId));
        if (!userDirectory.StartsWith(rootPrefix, StringComparison.Ordinal)
            || !Directory.Exists(userDirectory)
            || new DirectoryInfo(userDirectory).LinkTarget is not null)
        {
            throw new PptxValidationException("image_file_not_found", "The uploaded JPEG or PNG image was not found.");
        }

        var latest = string.Equals(fileId, "latest", StringComparison.OrdinalIgnoreCase);
        if (!latest && !SafeIdentifier().IsMatch(fileId))
        {
            throw InvalidFileId();
        }

        var candidates = Directory
            .EnumerateFiles(userDirectory, "*", SearchOption.TopDirectoryOnly)
            .Where(IsRegularImageUpload)
            .Where(path => latest || FileIdFromName(path).Equals(fileId, StringComparison.Ordinal))
            .OrderByDescending(path => new FileInfo(path).LastWriteTimeUtc)
            .ThenByDescending(Path.GetFileName, StringComparer.Ordinal)
            .Take(latest ? 1 : 2)
            .ToArray();
        if (candidates.Length != 1)
        {
            throw new PptxValidationException(
                candidates.Length == 0 ? "image_file_not_found" : "image_file_id_ambiguous",
                "Exactly one regular JPEG or PNG upload must match the opaque file identifier.");
        }

        var file = new FileInfo(candidates[0]);
        if (file.Length is <= 0 || file.Length > options.MaxImageFileBytes)
        {
            throw new PptxValidationException(
                "image_file_size_invalid",
                $"Uploaded images must be between 1 byte and {options.MaxImageFileBytes} bytes.");
        }

        var mediaType = await DetectMediaTypeAsync(file.FullName, cancellationToken).ConfigureAwait(false);
        var extensionMatches = mediaType == "image/png"
            ? file.Extension.Equals(".png", StringComparison.OrdinalIgnoreCase)
            : file.Extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase)
                || file.Extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase);
        if (!extensionMatches)
        {
            throw new PptxValidationException(
                "image_file_signature_invalid",
                "The upload extension must match its JPEG or PNG contents.");
        }
        return new ValidatedImageUpload(file.FullName, file.Length, FileIdFromName(file.FullName), mediaType);
    }

    private static bool IsRegularImageUpload(string path)
    {
        var file = new FileInfo(path);
        var extension = file.Extension.ToLowerInvariant();
        var delimiter = file.Name.IndexOf("__", StringComparison.Ordinal);
        return file.Exists
            && file.LinkTarget is null
            && extension is ".png" or ".jpg" or ".jpeg"
            && delimiter > 0
            && SafeIdentifier().IsMatch(file.Name[..delimiter]);
    }

    private static string FileIdFromName(string path) =>
        Path.GetFileName(path).Split("__", 2, StringSplitOptions.None)[0];

    private static async Task<string> DetectMediaTypeAsync(string path, CancellationToken cancellationToken)
    {
        var header = new byte[12];
        await using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        var bytesRead = await stream.ReadAsync(header, cancellationToken).ConfigureAwait(false);
        if (bytesRead >= PngSignature.Length && header.AsSpan(0, PngSignature.Length).SequenceEqual(PngSignature))
        {
            return "image/png";
        }

        if (bytesRead >= 3 && header[0] == 0xff && header[1] == 0xd8 && header[2] == 0xff)
        {
            return "image/jpeg";
        }

        throw new PptxValidationException(
            "image_file_signature_invalid",
            "The upload extension and contents must identify a valid JPEG or PNG image.");
    }

    private static PptxValidationException InvalidFileId() =>
        new("image_file_id_invalid", "The opaque image file identifier is invalid.");

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex SafeIdentifier();
}

public sealed record ValidatedImageUpload(string Path, long Bytes, string FileId, string MediaType);
