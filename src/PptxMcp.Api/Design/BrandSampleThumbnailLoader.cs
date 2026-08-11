using System.Buffers.Binary;
using System.Collections.ObjectModel;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using PptxMcp.Domain;

namespace PptxMcp.Design;

internal static class BrandSampleThumbnailLoader
{
    private const int MaximumThumbnailBytes = 256 * 1024;
    private const int MaximumProfileThumbnailBytes = 1024 * 1024;
    private const long MaximumProfileThumbnailPixels = 6_000_000;
    private const long MaximumProfileDecodedBytes = 32 * 1024 * 1024;
    private const int MaximumProfileThumbnails = 16;
    private const int MaximumWidth = 1_600;
    private const int MaximumHeight = 1_600;
    private const long MaximumPixels = 1_500_000;
    private static readonly byte[] PngSignature = [137, 80, 78, 71, 13, 10, 26, 10];

    public static IReadOnlyDictionary<string, BrandSampleThumbnail> Load(
        DirectoryInfo bundleDirectory,
        IReadOnlyList<BrandSampleSummary> samples)
    {
        ArgumentNullException.ThrowIfNull(bundleDirectory);
        ArgumentNullException.ThrowIfNull(samples);
        var thumbnailsDirectory = new DirectoryInfo(
            Path.Combine(bundleDirectory.FullName, "sample-thumbnails"));
        if (!thumbnailsDirectory.Exists)
        {
            return Empty();
        }

        if (thumbnailsDirectory.LinkTarget is not null)
        {
            throw new InvalidOperationException(
                "Brand Profile sample-thumbnails must not be a symbolic link.");
        }

        var loaded = new Dictionary<string, BrandSampleThumbnail>(StringComparer.Ordinal);
        var totalBytes = 0L;
        var totalPixels = 0L;
        var totalDecodedBytes = 0L;
        foreach (var sample in samples)
        {
            var unsupportedJpeg = new[]
            {
                new FileInfo(Path.Combine(thumbnailsDirectory.FullName, $"{sample.Id}.jpg")),
                new FileInfo(Path.Combine(thumbnailsDirectory.FullName, $"{sample.Id}.jpeg")),
            }.FirstOrDefault(static file => file.Exists);
            if (unsupportedJpeg is not null)
            {
                throw new InvalidOperationException(
                    $"Sample thumbnail {sample.Id} must be a metadata-free non-interlaced PNG; JPEG is not accepted.");
            }

            var file = new FileInfo(Path.Combine(thumbnailsDirectory.FullName, $"{sample.Id}.png"));
            if (!file.Exists)
            {
                continue;
            }

            if (file.LinkTarget is not null
                || file.Length is <= 0 or > MaximumThumbnailBytes)
            {
                throw new InvalidOperationException(
                    $"Sample thumbnail {sample.Id} must be a non-symbolic file of at most {MaximumThumbnailBytes} bytes.");
            }

            var bytes = ReadStableFile(file);
            totalBytes += bytes.Length;
            if (totalBytes > MaximumProfileThumbnailBytes)
            {
                throw new InvalidOperationException(
                    $"Brand Profile sample thumbnails must not exceed {MaximumProfileThumbnailBytes} bytes in total.");
            }

            var (mimeType, width, height, decodedBytes) = ValidateImage(bytes, sample.Id);
            if (width is <= 0 or > MaximumWidth
                || height is <= 0 or > MaximumHeight
                || (long)width * height > MaximumPixels)
            {
                throw new InvalidOperationException(
                    $"Sample thumbnail {sample.Id} dimensions exceed the allowed thumbnail bounds.");
            }

            totalPixels += (long)width * height;
            if (totalPixels > MaximumProfileThumbnailPixels)
            {
                throw new InvalidOperationException(
                    $"Brand Profile sample thumbnails must not exceed {MaximumProfileThumbnailPixels} pixels in total.");
            }

            totalDecodedBytes += decodedBytes;
            if (totalDecodedBytes > MaximumProfileDecodedBytes)
            {
                throw new InvalidOperationException(
                    $"Brand Profile sample thumbnails must not exceed {MaximumProfileDecodedBytes} decoded bytes in total.");
            }

            if (loaded.Count >= MaximumProfileThumbnails)
            {
                throw new InvalidOperationException(
                    $"Brand Profile must not contain more than {MaximumProfileThumbnails} sample thumbnails.");
            }

            loaded.Add(
                sample.Id,
                new BrandSampleThumbnail(sample.Id, mimeType, width, height, bytes, decodedBytes));
        }

        return new ReadOnlyDictionary<string, BrandSampleThumbnail>(loaded);
    }

    public static string ComputeContentHash(
        ReadOnlySpan<byte> manifestBytes,
        IReadOnlyDictionary<string, BrandSampleThumbnail> thumbnails)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        hash.AppendData(manifestBytes);
        ReadOnlySpan<byte> delimiter = [0];
        foreach (var pair in thumbnails.OrderBy(static pair => pair.Key, StringComparer.Ordinal))
        {
            hash.AppendData(delimiter);
            hash.AppendData(Encoding.UTF8.GetBytes(pair.Key));
            hash.AppendData(delimiter);
            hash.AppendData(Encoding.ASCII.GetBytes(pair.Value.MimeType));
            hash.AppendData(delimiter);
            hash.AppendData(SHA256.HashData(pair.Value.Bytes.Span));
        }

        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }

    private static byte[] ReadStableFile(FileInfo file)
    {
        byte[] bytes;
        using (var stream = new FileStream(
                   file.FullName,
                   FileMode.Open,
                   FileAccess.Read,
                   FileShare.Read,
                   bufferSize: 16 * 1024,
                   FileOptions.SequentialScan))
        {
            if (stream.Length is <= 0 or > MaximumThumbnailBytes)
            {
                throw new InvalidOperationException(
                    "Brand Profile sample thumbnail size changed during validation.");
            }

            bytes = new byte[checked((int)stream.Length)];
            stream.ReadExactly(bytes);
        }

        file.Refresh();
        if (!file.Exists || file.LinkTarget is not null || file.Length != bytes.Length)
        {
            throw new InvalidOperationException(
                "Brand Profile sample thumbnail changed during validation.");
        }

        return bytes;
    }

    private static (string MimeType, int Width, int Height, long DecodedBytes) ValidateImage(
        byte[] bytes,
        string sampleId)
    {
        if (bytes.Length < 24
            || !bytes.AsSpan(0, PngSignature.Length).SequenceEqual(PngSignature)
            || !bytes.AsSpan(12, 4).SequenceEqual("IHDR"u8))
        {
            throw new InvalidOperationException(
                $"Sample thumbnail {sampleId} MIME signature and content do not match PNG.");
        }

        var width = checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(16, 4)));
        var height = checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(20, 4)));
        if (width is <= 0 or > MaximumWidth
            || height is <= 0 or > MaximumHeight
            || (long)width * height > MaximumPixels)
        {
            throw new InvalidOperationException(
                $"Sample thumbnail {sampleId} dimensions exceed the allowed thumbnail bounds.");
        }

        var decodedBytes = ValidatePngChunks(bytes, sampleId, width, height);
        return ("image/png", width, height, decodedBytes);
    }

    private static long ValidatePngChunks(
        ReadOnlySpan<byte> bytes,
        string sampleId,
        int width,
        int height)
    {
        var offset = PngSignature.Length;
        var foundHeader = false;
        var foundImageData = false;
        var foundEnd = false;
        var imageDataEnded = false;
        var bitDepth = 0;
        var channelCount = 0;
        using var compressedImageData = new MemoryStream();
        while (offset + 12 <= bytes.Length)
        {
            var dataLength = BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(offset, 4));
            if (dataLength > int.MaxValue
                || offset + 12L + dataLength > bytes.Length)
            {
                throw new InvalidOperationException(
                    $"Sample thumbnail {sampleId} contains an invalid PNG chunk length.");
            }

            var type = bytes.Slice(offset + 4, 4);
            var allowed = type.SequenceEqual("IHDR"u8)
                || type.SequenceEqual("PLTE"u8)
                || type.SequenceEqual("IDAT"u8)
                || type.SequenceEqual("IEND"u8)
                || type.SequenceEqual("tRNS"u8)
                || type.SequenceEqual("sRGB"u8)
                || type.SequenceEqual("gAMA"u8)
                || type.SequenceEqual("cHRM"u8)
                || type.SequenceEqual("bKGD"u8);
            if (!allowed)
            {
                throw new InvalidOperationException(
                    $"Sample thumbnail {sampleId} contains a PNG metadata or unsupported chunk.");
            }

            var data = bytes.Slice(offset + 8, checked((int)dataLength));
            var expectedCrc = BinaryPrimitives.ReadUInt32BigEndian(
                bytes.Slice(offset + 8 + checked((int)dataLength), 4));
            if (ComputePngCrc(type, data) != expectedCrc)
            {
                throw new InvalidOperationException(
                    $"Sample thumbnail {sampleId} contains an invalid PNG checksum.");
            }

            if (type.SequenceEqual("IHDR"u8))
            {
                if (foundHeader || offset != PngSignature.Length || dataLength != 13)
                {
                    throw new InvalidOperationException($"Sample thumbnail {sampleId} has an invalid PNG header.");
                }

                foundHeader = true;
                bitDepth = data[8];
                channelCount = ValidatePngHeader(data, sampleId);
            }
            else if (type.SequenceEqual("IDAT"u8))
            {
                if (!foundHeader || imageDataEnded || dataLength == 0)
                {
                    throw new InvalidOperationException(
                        $"Sample thumbnail {sampleId} has invalid PNG image-data ordering.");
                }

                foundImageData = true;
                compressedImageData.Write(data);
            }
            else if (type.SequenceEqual("IEND"u8))
            {
                foundEnd = dataLength == 0 && offset + 12 == bytes.Length;
                break;
            }
            else if (foundImageData)
            {
                imageDataEnded = true;
            }

            offset += checked(12 + (int)dataLength);
        }

        if (!foundHeader || !foundImageData || !foundEnd)
        {
            throw new InvalidOperationException(
                $"Sample thumbnail {sampleId} is not a complete metadata-free PNG.");
        }

        ValidatePngImageData(
            compressedImageData.ToArray(),
            width,
            height,
            bitDepth,
            channelCount,
            sampleId);
        return checked(((long)width * channelCount * bitDepth + 7) / 8 + 1) * height;
    }

    private static int ValidatePngHeader(ReadOnlySpan<byte> data, string sampleId)
    {
        var bitDepth = data[8];
        var colorType = data[9];
        var validDepth = colorType switch
        {
            0 => bitDepth is 1 or 2 or 4 or 8 or 16,
            2 => bitDepth is 8 or 16,
            4 => bitDepth is 8 or 16,
            6 => bitDepth is 8 or 16,
            _ => false,
        };
        if (!validDepth
            || data[10] != 0
            || data[11] != 0
            || data[12] != 0)
        {
            throw new InvalidOperationException(
                $"Sample thumbnail {sampleId} must use a supported, non-interlaced PNG encoding.");
        }

        return colorType switch
        {
            0 => 1,
            2 => 3,
            4 => 2,
            6 => 4,
            _ => throw new InvalidOperationException($"Sample thumbnail {sampleId} has an invalid PNG color type."),
        };
    }

    private static void ValidatePngImageData(
        byte[] compressed,
        int width,
        int height,
        int bitDepth,
        int channelCount,
        string sampleId)
    {
        var rowBytes = checked((width * channelCount * bitDepth + 7) / 8);
        var scanline = new byte[checked(rowBytes + 1)];
        try
        {
            using var input = new MemoryStream(compressed, writable: false);
            using var inflater = new ZLibStream(input, CompressionMode.Decompress, leaveOpen: false);
            for (var row = 0; row < height; row++)
            {
                inflater.ReadExactly(scanline);
                if (scanline[0] > 4)
                {
                    throw new InvalidDataException("PNG scanline uses an invalid filter.");
                }
            }

            if (inflater.ReadByte() != -1)
            {
                throw new InvalidDataException("PNG image data exceeds the declared dimensions.");
            }
        }
        catch (Exception exception) when (exception is InvalidDataException or EndOfStreamException)
        {
            throw new InvalidOperationException(
                $"Sample thumbnail {sampleId} PNG image data could not be completely decoded.",
                exception);
        }
    }

    private static uint ComputePngCrc(ReadOnlySpan<byte> type, ReadOnlySpan<byte> data)
    {
        var crc = uint.MaxValue;
        foreach (var value in type)
        {
            crc = UpdatePngCrc(crc, value);
        }

        foreach (var value in data)
        {
            crc = UpdatePngCrc(crc, value);
        }

        return ~crc;
    }

    private static uint UpdatePngCrc(uint crc, byte value)
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

    private static ReadOnlyDictionary<string, BrandSampleThumbnail> Empty() =>
        new ReadOnlyDictionary<string, BrandSampleThumbnail>(
            new Dictionary<string, BrandSampleThumbnail>(StringComparer.Ordinal));
}
