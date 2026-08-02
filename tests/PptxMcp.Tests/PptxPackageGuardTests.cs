using System.IO.Compression;
using System.Text;
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

    [Fact]
    public async Task RejectsExternalRelationshipsBeforeOpeningPackage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                var entry = archive.CreateEntry("_rels/.rels");
                await using var stream = entry.Open();
                var xml = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"x\" Target=\"https://example.com\" TargetMode=\"External\"/></Relationships>";
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

    private static PptxPackageGuard CreateGuard() => new(Options.Create(new PptxMcpOptions
    {
        MaxFileBytes = 30 * 1024 * 1024,
        MaxSlides = 50,
    }));
}
