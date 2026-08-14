using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;

namespace PptxMcp.Storage;

public sealed partial class ImageAssetRepository(
    IOptions<PptxMcpOptions> options,
    UploadedImageResolver uploads,
    TimeProvider timeProvider)
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
    };
    private readonly PptxMcpOptions options = options.Value;
    private readonly string assetsRoot = Path.Combine(options.Value.StorageRoot, "image-assets");
    private readonly TimeSpan sanitizerTimeout = TimeSpan.FromSeconds(45);

    public string AssetsRoot => assetsRoot;

    public async Task<ImageAssetView> RegisterUserUploadAsync(
        CallerContext caller,
        string sourceFileId,
        string altText,
        string? attributionRef,
        CancellationToken cancellationToken)
    {
        ValidateAltText(altText);
        ValidateOptionalIdentifier(attributionRef, nameof(attributionRef));
        var source = await uploads.ResolveAsync(caller, sourceFileId, cancellationToken).ConfigureAwait(false);
        Directory.CreateDirectory(assetsRoot);

        var assetId = Guid.NewGuid().ToString("N");
        var directory = GetAssetDirectory(assetId);
        Directory.CreateDirectory(directory);
        var sanitizedPath = Path.Combine(directory, "asset.png");
        try
        {
            var sanitized = await SanitizeAsync(source.Path, sanitizedPath, cancellationToken).ConfigureAwait(false);
            var now = timeProvider.GetUtcNow();
            var manifest = new ImageAssetManifest(
                assetId,
                caller.UserScope,
                caller.ConversationScope,
                source.FileId,
                "userUpload",
                "userProvided",
                attributionRef,
                altText.Trim(),
                "image/png",
                sanitized.Width,
                sanitized.Height,
                sanitized.Bytes,
                await ComputeSha256Async(sanitizedPath, cancellationToken).ConfigureAwait(false),
                now,
                now.AddDays(options.RetentionDays));
            await WriteManifestAsync(directory, manifest, cancellationToken).ConfigureAwait(false);
            return CreateView(manifest);
        }
        catch
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }

            throw;
        }
    }

    public ImageAssetManifest GetOwned(CallerContext caller, string assetId)
    {
        var manifest = Get(assetId);
        if (!string.Equals(manifest.UserScope, caller.UserScope, StringComparison.Ordinal)
            || !string.Equals(manifest.ConversationScope, caller.ConversationScope, StringComparison.Ordinal)
            || manifest.ExpiresAt <= timeProvider.GetUtcNow())
        {
            throw NotFound();
        }

        return manifest;
    }

    public ImageAssetManifest Get(string assetId)
    {
        var directory = GetAssetDirectory(assetId);
        var manifestPath = Path.Combine(directory, "asset.json");
        var imagePath = Path.Combine(directory, "asset.png");
        if (!Directory.Exists(directory)
            || new DirectoryInfo(directory).LinkTarget is not null
            || !IsRegularFile(manifestPath)
            || !IsRegularFile(imagePath))
        {
            throw NotFound();
        }

        try
        {
            using var stream = File.Open(manifestPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var manifest = JsonSerializer.Deserialize<ImageAssetManifest>(stream, SerializerOptions);
            if (manifest is null
                || !string.Equals(manifest.AssetId, assetId, StringComparison.Ordinal)
                || manifest.MediaType != "image/png")
            {
                throw NotFound();
            }

            return manifest;
        }
        catch (JsonException)
        {
            throw NotFound();
        }
    }

    public IEnumerable<ImageAssetManifest> List()
    {
        if (!Directory.Exists(assetsRoot))
        {
            yield break;
        }

        foreach (var directory in Directory.EnumerateDirectories(assetsRoot))
        {
            var id = Path.GetFileName(directory);
            if (!AssetIdRegex().IsMatch(id))
            {
                continue;
            }

            ImageAssetManifest manifest;
            try
            {
                manifest = Get(id);
            }
            catch (PptxValidationException)
            {
                continue;
            }

            yield return manifest;
        }
    }

    public void Delete(string assetId)
    {
        var directory = GetAssetDirectory(assetId);
        if (Directory.Exists(directory) && new DirectoryInfo(directory).LinkTarget is null)
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private async Task<SanitizedImageResult> SanitizeAsync(
        string sourcePath,
        string destinationPath,
        CancellationToken cancellationToken)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "node",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            },
        };
        process.StartInfo.ArgumentList.Add(options.ImageSanitizerPath);
        process.StartInfo.ArgumentList.Add(sourcePath);
        process.StartInfo.ArgumentList.Add(destinationPath);
        process.StartInfo.ArgumentList.Add(options.MaxImagePixels.ToString(System.Globalization.CultureInfo.InvariantCulture));
        process.StartInfo.ArgumentList.Add(options.MaxImageDimension.ToString(System.Globalization.CultureInfo.InvariantCulture));
        process.StartInfo.ArgumentList.Add(options.MaxImageFileBytes.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (!process.Start())
        {
            throw SanitizationFailed("The image sanitizer could not be started.");
        }

        var outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var errorTask = process.StandardError.ReadToEndAsync(cancellationToken);
        using var timeout = new CancellationTokenSource(sanitizerTimeout);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeout.Token);
        try
        {
            await process.WaitForExitAsync(linked.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }

            throw;
        }

        var output = await outputTask.ConfigureAwait(false);
        await errorTask.ConfigureAwait(false);
        if (process.ExitCode != 0 || !IsRegularFile(destinationPath))
        {
            throw SanitizationFailed();
        }

        try
        {
            var result = JsonSerializer.Deserialize<SanitizedImageResult>(output, SerializerOptions);
            if (result is null
                || result.Width is <= 0 or > 4_096
                || result.Height is <= 0 or > 4_096
                || result.Bytes is <= 0
                || result.Bytes != new FileInfo(destinationPath).Length)
            {
                throw SanitizationFailed("The sanitizer returned inconsistent image metadata.");
            }

            return result;
        }
        catch (JsonException)
        {
            throw SanitizationFailed("The sanitizer returned invalid metadata.");
        }
    }

    private string GetAssetDirectory(string assetId)
    {
        if (!AssetIdRegex().IsMatch(assetId))
        {
            throw NotFound();
        }

        var root = Path.GetFullPath(assetsRoot);
        var directory = Path.GetFullPath(Path.Combine(root, assetId));
        if (!directory.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal))
        {
            throw NotFound();
        }

        return directory;
    }

    private static async Task WriteManifestAsync(
        string directory,
        ImageAssetManifest manifest,
        CancellationToken cancellationToken)
    {
        var path = Path.Combine(directory, "asset.json");
        var temporaryPath = Path.Combine(directory, $"asset-{Guid.NewGuid():N}.tmp");
        await using (var stream = File.Open(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            await JsonSerializer.SerializeAsync(stream, manifest, SerializerOptions, cancellationToken).ConfigureAwait(false);
        }

        File.Move(temporaryPath, path);
    }

    private static async Task<string> ComputeSha256Async(string path, CancellationToken cancellationToken)
    {
        await using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(false)).ToLowerInvariant();
    }

    private static bool IsRegularFile(string path)
    {
        var file = new FileInfo(path);
        return file.Exists && file.LinkTarget is null;
    }

    private static void ValidateAltText(string value)
    {
        if (string.IsNullOrWhiteSpace(value)
            || value.Length > 240
            || value.Any(static character => char.IsControl(character)))
        {
            throw new PptxValidationException(
                "image_asset_alt_text_invalid",
                "altText must be a concise 1-240 character description without control characters.");
        }
    }

    private static void ValidateOptionalIdentifier(string? value, string field)
    {
        if (value is not null && !OpaqueIdentifierRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "image_asset_attribution_invalid",
                $"{field} must be an opaque identifier, not source text, a URL, or a path.");
        }
    }

    private static PptxValidationException NotFound() =>
        new("image_asset_not_found", "The image asset was not found for this user and conversation, or it expired.");

    private static PptxValidationException SanitizationFailed(string? reason = null) =>
        new(
            "image_asset_sanitization_failed",
            reason is null
                ? "The uploaded image could not be safely normalized. Use a valid single-frame JPEG or PNG within the published limits."
                : reason);

    private static ImageAssetView CreateView(ImageAssetManifest manifest) =>
        new(
            manifest.AssetId,
            "ready",
            manifest.MediaType,
            manifest.Width,
            manifest.Height,
            manifest.AltText,
            manifest.Acquisition,
            manifest.LicenseStatus,
            manifest.AttributionRef,
            manifest.ExpiresAt,
            "Use this asset_id only in the current conversation's Asset Plan and Media slide. Do not expose or transform it into a path or URL.");

    [GeneratedRegex("\\A[0-9a-f]{32}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex AssetIdRegex();

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex OpaqueIdentifierRegex();
}

public sealed record ImageAssetManifest(
    string AssetId,
    string UserScope,
    string ConversationScope,
    string SourceFileId,
    string Acquisition,
    string LicenseStatus,
    string? AttributionRef,
    string AltText,
    string MediaType,
    int Width,
    int Height,
    long Bytes,
    string Sha256,
    DateTimeOffset CreatedAt,
    DateTimeOffset ExpiresAt);

public sealed record ImageAssetView(
    [property: System.Text.Json.Serialization.JsonPropertyName("asset_id")] string AssetId,
    [property: System.Text.Json.Serialization.JsonPropertyName("status")] string Status,
    [property: System.Text.Json.Serialization.JsonPropertyName("media_type")] string MediaType,
    [property: System.Text.Json.Serialization.JsonPropertyName("width")] int Width,
    [property: System.Text.Json.Serialization.JsonPropertyName("height")] int Height,
    [property: System.Text.Json.Serialization.JsonPropertyName("alt_text")] string AltText,
    [property: System.Text.Json.Serialization.JsonPropertyName("acquisition")] string Acquisition,
    [property: System.Text.Json.Serialization.JsonPropertyName("license_status")] string LicenseStatus,
    [property: System.Text.Json.Serialization.JsonPropertyName("attribution_ref")] string? AttributionRef,
    [property: System.Text.Json.Serialization.JsonPropertyName("expires_at")] DateTimeOffset ExpiresAt,
    [property: System.Text.Json.Serialization.JsonPropertyName("instruction")] string Instruction);

internal sealed record SanitizedImageResult(int Width, int Height, long Bytes);
