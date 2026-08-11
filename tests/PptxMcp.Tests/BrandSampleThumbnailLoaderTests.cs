using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using PptxMcp.Design;
using PptxMcp.Domain;

namespace PptxMcp.Tests;

public sealed class BrandSampleThumbnailLoaderTests
{
    [Fact]
    public void LoadsStrictPngAndBindsThumbnailBytesIntoProfileHash()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        try
        {
            var bundle = CreateBundle(root);
            var samples = new[] { Sample("sample-one") };
            var firstBytes = CreatePng(2, 1, 8, 2, 0x11);
            File.WriteAllBytes(Path.Combine(bundle.FullName, "sample-thumbnails", "sample-one.png"), firstBytes);

            var first = BrandSampleThumbnailLoader.Load(bundle, samples);
            var firstHash = BrandSampleThumbnailLoader.ComputeContentHash("manifest"u8, first);
            var firstThumbnail = Assert.Single(first).Value;
            Assert.Equal("image/png", firstThumbnail.MimeType);
            Assert.Equal(2, firstThumbnail.Width);
            Assert.Equal(1, firstThumbnail.Height);
            Assert.True(firstThumbnail.DecodedBytes > 0);

            var secondBytes = CreatePng(2, 1, 8, 2, 0x22);
            File.WriteAllBytes(Path.Combine(bundle.FullName, "sample-thumbnails", "sample-one.png"), secondBytes);
            var second = BrandSampleThumbnailLoader.Load(bundle, samples);
            var secondHash = BrandSampleThumbnailLoader.ComputeContentHash("manifest"u8, second);

            Assert.NotEqual(firstHash, secondHash);
            Assert.NotEqual(
                Convert.ToHexString(System.Security.Cryptography.SHA256.HashData("manifest"u8)).ToLowerInvariant(),
                firstHash);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsPngMetadataInvalidCrcJpegAndSymbolicLinks()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        try
        {
            var bundle = CreateBundle(root);
            var directory = Path.Combine(bundle.FullName, "sample-thumbnails");
            var path = Path.Combine(directory, "sample-one.png");
            File.WriteAllBytes(path, CreatePng(2, 1, 8, 2, 0x11, includeTextChunk: true));
            Assert.Contains(
                "metadata or unsupported chunk",
                Assert.Throws<InvalidOperationException>(() =>
                    BrandSampleThumbnailLoader.Load(bundle, [Sample("sample-one")])).Message,
                StringComparison.Ordinal);

            var invalidCrc = CreatePng(2, 1, 8, 2, 0x11);
            invalidCrc[^1] ^= 0x01;
            File.WriteAllBytes(path, invalidCrc);
            Assert.Contains(
                "checksum",
                Assert.Throws<InvalidOperationException>(() =>
                    BrandSampleThumbnailLoader.Load(bundle, [Sample("sample-one")])).Message,
                StringComparison.Ordinal);

            File.Delete(path);
            File.WriteAllBytes(Path.Combine(directory, "sample-one.jpg"), [0xff, 0xd8, 0xff, 0xd9]);
            Assert.Contains(
                "JPEG is not accepted",
                Assert.Throws<InvalidOperationException>(() =>
                    BrandSampleThumbnailLoader.Load(bundle, [Sample("sample-one")])).Message,
                StringComparison.Ordinal);
            File.Delete(Path.Combine(directory, "sample-one.jpg"));

            var target = Path.Combine(bundle.FullName, "target.png");
            File.WriteAllBytes(target, CreatePng(1, 1, 8, 2, 0x11));
            File.CreateSymbolicLink(path, target);
            Assert.Contains(
                "non-symbolic",
                Assert.Throws<InvalidOperationException>(() =>
                    BrandSampleThumbnailLoader.Load(bundle, [Sample("sample-one")])).Message,
                StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsDimensionAndAggregateDecodedByteBombs()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        try
        {
            var bundle = CreateBundle(root);
            var directory = Path.Combine(bundle.FullName, "sample-thumbnails");
            File.WriteAllBytes(
                Path.Combine(directory, "oversized.png"),
                CreatePng(1_601, 1, 8, 2, 0x00));
            Assert.Contains(
                "dimensions",
                Assert.Throws<InvalidOperationException>(() =>
                    BrandSampleThumbnailLoader.Load(bundle, [Sample("oversized")])).Message,
                StringComparison.Ordinal);

            File.Delete(Path.Combine(directory, "oversized.png"));
            var samples = Enumerable.Range(1, 3)
                .Select(index => Sample($"large-{index}"))
                .ToArray();
            foreach (var sample in samples)
            {
                File.WriteAllBytes(
                    Path.Combine(directory, $"{sample.Id}.png"),
                    CreatePng(1_500, 1_000, 16, 6, 0x00));
            }

            var error = Assert.Throws<InvalidOperationException>(() =>
                BrandSampleThumbnailLoader.Load(bundle, samples));
            Assert.Contains("decoded bytes", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsPerProfileThumbnailCountAndAggregatePixelCaps()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        try
        {
            var countBundle = CreateBundle(Path.Combine(root, "count"));
            var countSamples = Enumerable.Range(1, 17)
                .Select(index => Sample($"count-{index}"))
                .ToArray();
            foreach (var sample in countSamples)
            {
                File.WriteAllBytes(
                    Path.Combine(countBundle.FullName, "sample-thumbnails", $"{sample.Id}.png"),
                    CreatePng(1, 1, 8, 2, 0x00));
            }

            Assert.Contains(
                "more than 16",
                Assert.Throws<InvalidOperationException>(() =>
                    BrandSampleThumbnailLoader.Load(countBundle, countSamples)).Message,
                StringComparison.Ordinal);

            var pixelBundle = CreateBundle(Path.Combine(root, "pixels"));
            var pixelSamples = Enumerable.Range(1, 5)
                .Select(index => Sample($"pixels-{index}"))
                .ToArray();
            foreach (var sample in pixelSamples)
            {
                File.WriteAllBytes(
                    Path.Combine(pixelBundle.FullName, "sample-thumbnails", $"{sample.Id}.png"),
                    CreatePng(1_300, 1_000, 8, 2, 0x00));
            }

            Assert.Contains(
                "pixels in total",
                Assert.Throws<InvalidOperationException>(() =>
                    BrandSampleThumbnailLoader.Load(pixelBundle, pixelSamples)).Message,
                StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void CatalogRejectsAggregateDecodedThumbnailWorkAcrossProfiles()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        try
        {
            var manifestTemplate = JsonNode.Parse(
                File.ReadAllText(Path.Combine(root, "general", "brand-profile.json")))!.AsObject();
            foreach (var profileId in new[] { "general", "profile-a", "profile-b", "profile-c", "profile-d", "profile-e" })
            {
                WriteLargeThumbnailProfile(root, profileId, manifestTemplate);
            }

            var catalog = BrandProfileTestFactory.CreateCatalog(root);
            var error = Assert.Throws<InvalidOperationException>(catalog.EnsureReady);

            Assert.Contains("catalog sample thumbnails", error.Message, StringComparison.Ordinal);
            Assert.Contains("decoded bytes", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private static DirectoryInfo CreateBundle(string root)
    {
        var bundle = Directory.CreateDirectory(Path.Combine(root, "bundle"));
        Directory.CreateDirectory(Path.Combine(bundle.FullName, "sample-thumbnails"));
        return bundle;
    }

    private static void WriteLargeThumbnailProfile(
        string root,
        string profileId,
        JsonObject manifestTemplate)
    {
        var manifest = manifestTemplate.DeepClone().AsObject();
        manifest["id"] = profileId;
        var recipes = new JsonArray();
        var samples = new JsonArray();
        for (var index = 1; index <= 4; index++)
        {
            var recipeId = $"cover-{index}";
            var sampleId = $"sample-{index}";
            recipes.Add(JsonNode.Parse($$"""{"id":"{{recipeId}}","purpose":"cover","semantic_kind":"Title","variant":"auto","density":"airy","style_direction_id":"standard","required_asset_roles":[],"sample_ids":["{{sampleId}}"]}"""));
            samples.Add(JsonNode.Parse($$"""{"id":"{{sampleId}}","title":"Sample {{index}}","summary":"Approved sample.","purpose":"cover","density":"airy","style_direction_id":"standard","recipe_id":"{{recipeId}}","information_level":"low"}"""));
        }

        manifest["layout_recipes"] = recipes;
        manifest["samples"] = samples;
        var bundle = Directory.CreateDirectory(Path.Combine(root, profileId));
        File.WriteAllText(Path.Combine(bundle.FullName, "brand-profile.json"), manifest.ToJsonString());
        var thumbnails = Directory.CreateDirectory(Path.Combine(bundle.FullName, "sample-thumbnails"));
        for (var index = 1; index <= 4; index++)
        {
            File.WriteAllBytes(
                Path.Combine(thumbnails.FullName, $"sample-{index}.png"),
                CreatePng(1_500, 1_000, 8, 6, 0x00));
        }
    }

    private static BrandSampleSummary Sample(string id) => new(
        id,
        id,
        "Sample summary",
        "cover",
        "airy",
        "standard",
        "cover-airy",
        "low");

    internal static byte[] CreatePng(
        int width,
        int height,
        byte bitDepth,
        byte colorType,
        byte pixelValue,
        bool includeTextChunk = false)
    {
        var channels = colorType switch
        {
            0 => 1,
            2 => 3,
            4 => 2,
            6 => 4,
            _ => throw new ArgumentOutOfRangeException(nameof(colorType)),
        };
        var rowBytes = checked((width * channels * bitDepth + 7) / 8);
        var raw = new byte[checked((rowBytes + 1) * height)];
        for (var row = 0; row < height; row++)
        {
            raw[row * (rowBytes + 1)] = 0;
            raw.AsSpan(row * (rowBytes + 1) + 1, rowBytes).Fill(pixelValue);
        }

        byte[] compressed;
        using (var output = new MemoryStream())
        {
            using (var compressor = new ZLibStream(output, CompressionLevel.SmallestSize, leaveOpen: true))
            {
                compressor.Write(raw);
            }

            compressed = output.ToArray();
        }

        using var png = new MemoryStream();
        png.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        var header = new byte[13];
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(0, 4), checked((uint)width));
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4, 4), checked((uint)height));
        header[8] = bitDepth;
        header[9] = colorType;
        WriteChunk(png, "IHDR"u8, header);
        if (includeTextChunk)
        {
            WriteChunk(png, "tEXt"u8, "comment\0secret"u8);
        }

        WriteChunk(png, "IDAT"u8, compressed);
        WriteChunk(png, "IEND"u8, []);
        return png.ToArray();
    }

    private static void WriteChunk(Stream output, ReadOnlySpan<byte> type, ReadOnlySpan<byte> data)
    {
        Span<byte> length = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(length, checked((uint)data.Length));
        output.Write(length);
        output.Write(type);
        output.Write(data);
        Span<byte> crc = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crc, ComputeCrc(type, data));
        output.Write(crc);
    }

    private static uint ComputeCrc(ReadOnlySpan<byte> type, ReadOnlySpan<byte> data)
    {
        var crc = uint.MaxValue;
        foreach (var value in type)
        {
            crc = UpdateCrc(crc, value);
        }

        foreach (var value in data)
        {
            crc = UpdateCrc(crc, value);
        }

        return ~crc;
    }

    private static uint UpdateCrc(uint crc, byte value)
    {
        crc ^= value;
        for (var bit = 0; bit < 8; bit++)
        {
            crc = (crc & 1) == 0
                ? crc >> 1
                : 0xedb88320U ^ (crc >> 1);
        }

        return crc;
    }
}
