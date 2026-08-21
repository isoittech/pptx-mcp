using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class PptxPackageGuardTests
{
    [Fact]
    public async Task AcceptsPresentationWithinLimits()
    {
        var path = TestPresentationFactory.Create("Hello");
        try
        {
            var result = await CreateGuard().ValidateAsync(path, CancellationToken.None);
            Assert.Equal(1, result.SlideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("https://example.com/resource")]
    [InlineData("http://example.com/resource")]
    public async Task AcceptsOrdinaryHttpAndHttpsHyperlinkRelationships(string target)
    {
        var path = TestPresentationFactory.Create("Web link");
        try
        {
            using (var document = PresentationDocument.Open(path, true))
            {
                var slide = Assert.Single(document.PresentationPart!.SlideParts);
                slide.AddHyperlinkRelationship(new Uri(target), true);
            }

            var result = await CreateGuard().ValidateAsync(path, CancellationToken.None);

            Assert.Equal(1, result.SlideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "https://files.example.com/image.png")]
    [InlineData("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", "file:///shared/source.xlsx")]
    [InlineData("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", "file://server/share/template.potx")]
    public async Task RejectsExternalImageOleAndSharedFileRelationships(
        string relationshipType,
        string target)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                var entry = archive.CreateEntry("_rels/.rels");
                await using var stream = entry.Open();
                var xml = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{relationshipType}\" Target=\"{target}\" TargetMode=\"External\"/></Relationships>";
                await stream.WriteAsync(Encoding.UTF8.GetBytes(xml), CancellationToken.None);
            }

            var exception = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard().ValidateAsync(path, CancellationToken.None));
            Assert.Equal("external_relationship", exception.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RejectsNonWebHyperlinkRelationship()
    {
        var path = TestPresentationFactory.Create("Local link");
        try
        {
            using (var document = PresentationDocument.Open(path, true))
            {
                var slide = Assert.Single(document.PresentationPart!.SlideParts);
                slide.AddHyperlinkRelationship(new Uri("file:///etc/passwd"), true);
            }

            var exception = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard().ValidateAsync(path, CancellationToken.None));

            Assert.Equal("external_relationship", exception.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RejectsMacroEnabledContentType()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                var entry = archive.CreateEntry("[Content_Types].xml");
                await using var stream = entry.Open();
                var xml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml\"/></Types>";
                await stream.WriteAsync(Encoding.UTF8.GetBytes(xml), CancellationToken.None);
            }

            var exception = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard().ValidateAsync(path, CancellationToken.None));
            Assert.Equal("active_content", exception.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RejectsActiveXPackageEntry()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                var entry = archive.CreateEntry("ppt/activeX/activeX1.xml");
                await using var stream = entry.Open();
                await stream.WriteAsync(Encoding.UTF8.GetBytes("<activeX />"), CancellationToken.None);
            }

            var exception = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard().ValidateAsync(path, CancellationToken.None));

            Assert.Equal("active_content", exception.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RejectsUnreadableZipWithStructuredCode()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            await File.WriteAllTextAsync(path, "This is not a ZIP package.");

            var exception = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard().ValidateAsync(path, CancellationToken.None));

            Assert.Equal("invalid_zip", exception.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RejectsZipThatIsNotAPresentationWithStructuredCode()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                var entry = archive.CreateEntry("[Content_Types].xml");
                await using var stream = entry.Open();
                const string xml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>";
                await stream.WriteAsync(Encoding.UTF8.GetBytes(xml), CancellationToken.None);
            }

            var exception = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard().ValidateAsync(path, CancellationToken.None));

            Assert.Equal("invalid_pptx", exception.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RejectsEmptyAndOversizedFilesWithStructuredCode()
    {
        var emptyPath = Path.Combine(Path.GetTempPath(), $"pptx-mcp-empty-{Guid.NewGuid():N}.pptx");
        var oversizedPath = Path.Combine(Path.GetTempPath(), $"pptx-mcp-large-{Guid.NewGuid():N}.pptx");
        try
        {
            await File.WriteAllBytesAsync(emptyPath, []);
            await File.WriteAllBytesAsync(oversizedPath, new byte[11]);

            var empty = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard(maxFileBytes: 10).ValidateAsync(emptyPath, CancellationToken.None));
            var oversized = await Assert.ThrowsAsync<PptxValidationException>(
                () => CreateGuard(maxFileBytes: 10).ValidateAsync(oversizedPath, CancellationToken.None));

            Assert.Equal("file_size_out_of_range", empty.Code);
            Assert.Equal("file_size_out_of_range", oversized.Code);
        }
        finally
        {
            File.Delete(emptyPath);
            File.Delete(oversizedPath);
        }
    }

    private static PptxPackageGuard CreateGuard(long? maxFileBytes = null) => new(Options.Create(new PptxMcpOptions
    {
        MaxFileBytes = maxFileBytes ?? 30 * 1024 * 1024,
        MaxSlides = 50,
    }));
}
