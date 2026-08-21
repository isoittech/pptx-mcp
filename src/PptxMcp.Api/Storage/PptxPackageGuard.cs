using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;

namespace PptxMcp.Storage;

public sealed class PptxPackageGuard(IOptions<PptxMcpOptions> options)
{
    private const string TransitionalHyperlinkRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private const string StrictHyperlinkRelationship =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink";
    private readonly PptxMcpOptions options = options.Value;

    public Task<ValidatedInput> ValidateAsync(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var file = new FileInfo(path);
        if (!file.Exists || file.Length == 0 || file.Length > options.MaxFileBytes)
        {
            throw new PptxValidationException(
                "file_size_out_of_range",
                $"PPTX files must be between 1 byte and {options.MaxFileBytes} bytes.");
        }

        ValidateZipEnvelope(path, cancellationToken);

        int slideCount;
        try
        {
            using var document = PresentationDocument.Open(path, false);
            var presentationPart = document.PresentationPart
                ?? throw new PptxValidationException(
                    "invalid_pptx",
                    "The PPTX package does not contain a presentation part.");
            var presentation = presentationPart.Presentation
                ?? throw new PptxValidationException(
                    "invalid_pptx",
                    "The PPTX package does not contain a presentation root.");
            slideCount = presentation?.SlideIdList?.ChildElements.Count ?? 0;
        }
        catch (PptxValidationException)
        {
            throw;
        }
        catch (Exception exception) when (exception is OpenXmlPackageException
            or FileFormatException
            or InvalidDataException
            or XmlException)
        {
            throw new PptxValidationException("invalid_pptx", $"The PPTX package is invalid: {exception.Message}");
        }

        if (slideCount is <= 0 || slideCount > 50 || slideCount > options.MaxSlides)
        {
            throw new PptxValidationException(
                "slide_count_out_of_range",
                $"PPTX files must contain between 1 and {options.MaxSlides} slides.");
        }

        return Task.FromResult(new ValidatedInput(path, file.Length, slideCount));
    }

    private void ValidateZipEnvelope(string path, CancellationToken cancellationToken)
    {
        try
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);

            if (archive.Entries.Count > options.MaxZipEntries)
            {
                throw new PptxValidationException("zip_entry_limit", "The PPTX contains too many ZIP entries.");
            }

            long totalLength = 0;
            foreach (var entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidateEntryName(entry.FullName);
                totalLength = checked(totalLength + entry.Length);

                if (totalLength > options.MaxUncompressedBytes)
                {
                    throw new PptxValidationException("zip_expansion_limit", "The uncompressed PPTX is too large.");
                }

                if (entry.Length > 1_048_576
                    && entry.CompressedLength > 0
                    && entry.Length / entry.CompressedLength > options.MaxCompressionRatio)
                {
                    throw new PptxValidationException("zip_compression_ratio", "A PPTX entry has an unsafe compression ratio.");
                }

                var normalized = entry.FullName.Replace('\\', '/');
                if (normalized.Contains("/vbaProject", StringComparison.OrdinalIgnoreCase)
                    || normalized.Contains("/activeX/", StringComparison.OrdinalIgnoreCase))
                {
                    throw new PptxValidationException("active_content", "Macro and ActiveX content is not accepted.");
                }

                if (string.Equals(normalized, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                {
                    RejectActiveContentTypes(entry);
                }

                if (normalized.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    RejectExternalRelationships(entry);
                }
            }
        }
        catch (InvalidDataException exception)
        {
            throw new PptxValidationException("invalid_zip", $"The file is not a readable PPTX package: {exception.Message}");
        }
        catch (OverflowException)
        {
            throw new PptxValidationException("zip_expansion_limit", "The uncompressed PPTX is too large.");
        }
        catch (XmlException exception)
        {
            throw new PptxValidationException("invalid_xml", $"The PPTX contains invalid package XML: {exception.Message}");
        }
    }

    private static void ValidateEntryName(string entryName)
    {
        var normalized = entryName.Replace('\\', '/');
        if (normalized.StartsWith('/')
            || normalized.Split('/').Any(segment => segment is ".." or "."))
        {
            throw new PptxValidationException("zip_path_traversal", "The PPTX contains an unsafe ZIP entry name.");
        }
    }

    private static void RejectExternalRelationships(ZipArchiveEntry entry)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 4 * 1024 * 1024,
        };

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, settings);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element
                && string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal)
                && string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)
                && !IsSafeWebHyperlink(reader))
            {
                throw new PptxValidationException(
                    "external_relationship",
                    "PPTX files with external resources or non-web hyperlinks are not accepted.");
            }
        }
    }

    private static bool IsSafeWebHyperlink(XmlReader reader)
    {
        var relationshipType = reader.GetAttribute("Type");
        if (!string.Equals(relationshipType, TransitionalHyperlinkRelationship, StringComparison.Ordinal)
            && !string.Equals(relationshipType, StrictHyperlinkRelationship, StringComparison.Ordinal))
        {
            return false;
        }

        return Uri.TryCreate(reader.GetAttribute("Target"), UriKind.Absolute, out var target)
            && (target.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || target.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase));
    }

    private static void RejectActiveContentTypes(ZipArchiveEntry entry)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 4 * 1024 * 1024,
        };

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, settings);
        while (reader.Read())
        {
            var contentType = reader.GetAttribute("ContentType");
            if (contentType?.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase) == true
                || contentType?.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) == true)
            {
                throw new PptxValidationException("active_content", "Macro-enabled PowerPoint content is not accepted.");
            }
        }
    }
}
