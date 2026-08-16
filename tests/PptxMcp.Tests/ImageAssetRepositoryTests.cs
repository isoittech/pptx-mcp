using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class ImageAssetRepositoryTests
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);
    [Fact]
    public async Task UploadedImageResolverAcceptsOpaqueJpegOrPngIdsAndRejectsSpoofedMagic()
    {
        var root = Path.Combine(Path.GetTempPath(), $"pptx-image-upload-{Guid.NewGuid():N}");
        var uploadsRoot = Path.Combine(root, "uploads");
        var userDirectory = Path.Combine(uploadsRoot, "user-a");
        Directory.CreateDirectory(userDirectory);
        var options = Options.Create(new PptxMcpOptions
        {
            LibreChatUploadsRoot = uploadsRoot,
            StorageRoot = Path.Combine(root, "storage"),
        });
        var resolver = new UploadedImageResolver(options);
        var caller = new CallerContext("user-a", "conversation-a", null);

        try
        {
            var pngPath = Path.Combine(userDirectory, "image123__approved.png");
            await File.WriteAllBytesAsync(pngPath, [137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0]);

            var resolved = await resolver.ResolveAsync(caller, "image123", CancellationToken.None);

            Assert.Equal("image123", resolved.FileId);
            Assert.Equal("image/png", resolved.MediaType);

            var spoofPath = Path.Combine(userDirectory, "image456__spoof.png");
            await File.WriteAllTextAsync(spoofPath, "not an image");
            var error = await Assert.ThrowsAsync<PptxValidationException>(() =>
                resolver.ResolveAsync(caller, "image456", CancellationToken.None));
            Assert.Equal("image_file_signature_invalid", error.Code);

            var jpegAsPngPath = Path.Combine(userDirectory, "image789__wrong-extension.png");
            await File.WriteAllBytesAsync(jpegAsPngPath, [0xff, 0xd8, 0xff, 0, 0, 0]);
            error = await Assert.ThrowsAsync<PptxValidationException>(() =>
                resolver.ResolveAsync(caller, "image789", CancellationToken.None));
            Assert.Equal("image_file_signature_invalid", error.Code);

            var symlinkPath = Path.Combine(userDirectory, "image999__linked.png");
            File.CreateSymbolicLink(symlinkPath, pngPath);
            error = await Assert.ThrowsAsync<PptxValidationException>(() =>
                resolver.ResolveAsync(caller, "image999", CancellationToken.None));
            Assert.Equal("image_file_not_found", error.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task UploadedImageResolverUsesLibreChatMessageImageRoot()
    {
        var root = Path.Combine(Path.GetTempPath(), $"pptx-message-image-{Guid.NewGuid():N}");
        var imagesRoot = Path.Combine(root, "images");
        var userDirectory = Path.Combine(imagesRoot, "user-a");
        Directory.CreateDirectory(userDirectory);
        var options = Options.Create(new PptxMcpOptions
        {
            LibreChatImagesRoot = imagesRoot,
            LibreChatUploadsRoot = Path.Combine(root, "uploads"),
            StorageRoot = Path.Combine(root, "storage"),
        });
        var resolver = new UploadedImageResolver(options);
        var caller = new CallerContext("user-a", "conversation-a", null);
        const string fileId = "81a85f9b-cf73-4950-ac25-c63c96694473";

        try
        {
            var imagePath = Path.Combine(userDirectory, $"{fileId}__subtle-upload.png");
            await File.WriteAllBytesAsync(imagePath, [137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0]);

            var explicitImage = await resolver.ResolveAsync(caller, fileId, CancellationToken.None);
            var latestImage = await resolver.ResolveAsync(caller, "latest", CancellationToken.None);

            Assert.Equal(fileId, explicitImage.FileId);
            Assert.Equal(imagePath, explicitImage.Path);
            Assert.Equal(fileId, latestImage.FileId);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task StoredAssetIsBoundToUserConversationAndExpiry()
    {
        var root = Path.Combine(Path.GetTempPath(), $"pptx-image-asset-{Guid.NewGuid():N}");
        var options = Options.Create(new PptxMcpOptions
        {
            LibreChatUploadsRoot = Path.Combine(root, "uploads"),
            StorageRoot = Path.Combine(root, "storage"),
            RetentionDays = 7,
        });
        var clock = new TestTimeProvider(new DateTimeOffset(2026, 8, 14, 0, 0, 0, TimeSpan.Zero));
        var repository = new ImageAssetRepository(options, new UploadedImageResolver(options), clock);
        var owner = new CallerContext("user-a", "conversation-a", null);
        var assetId = "0123456789abcdef0123456789abcdef";
        var assetDirectory = Path.Combine(repository.AssetsRoot, assetId);
        Directory.CreateDirectory(assetDirectory);
        var manifest = new ImageAssetManifest(
            assetId,
            owner.UserScope,
            owner.ConversationScope,
            "source-file",
            "userUpload",
            "userProvided",
            "ATTR-001",
            "A product in use",
            "image/png",
            800,
            450,
            12,
            new string('a', 64),
            clock.GetUtcNow(),
            clock.GetUtcNow().AddDays(7));
        await File.WriteAllBytesAsync(Path.Combine(assetDirectory, "asset.png"), [137, 80, 78, 71, 13, 10, 26, 10]);
        await File.WriteAllTextAsync(
            Path.Combine(assetDirectory, "asset.json"),
            JsonSerializer.Serialize(manifest, SerializerOptions));

        try
        {
            Assert.Equal(assetId, repository.GetOwned(owner, assetId).AssetId);
            Assert.Equal(
                "image_asset_not_found",
                Assert.Throws<PptxValidationException>(() => repository.GetOwned(
                    new CallerContext("user-b", "conversation-a", null),
                    assetId)).Code);
            Assert.Equal(
                "image_asset_not_found",
                Assert.Throws<PptxValidationException>(() => repository.GetOwned(
                    new CallerContext("user-a", "conversation-b", null),
                    assetId)).Code);

            clock.Advance(TimeSpan.FromDays(8));
            Assert.Equal(
                "image_asset_not_found",
                Assert.Throws<PptxValidationException>(() => repository.GetOwned(owner, assetId)).Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private sealed class TestTimeProvider(DateTimeOffset now) : TimeProvider
    {
        private DateTimeOffset now = now;

        public override DateTimeOffset GetUtcNow() => now;

        public void Advance(TimeSpan duration) => now = now.Add(duration);
    }
}
