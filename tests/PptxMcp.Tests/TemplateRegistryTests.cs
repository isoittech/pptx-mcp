using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class TemplateRegistryTests
{
    [Fact]
    public async Task ResolvesValidatedDefaultTemplateByIdentifier()
    {
        var templatesRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-templates-{Guid.NewGuid():N}");
        Directory.CreateDirectory(templatesRoot);
        var expectedPath = Path.Combine(templatesRoot, "organization-default.pptx");
        File.Move(TestPresentationFactory.CreateBlankBrandedTemplate(), expectedPath);

        try
        {
            var registry = CreateRegistry(templatesRoot, "organization-default");

            var result = await registry.ResolveDefaultAsync(CancellationToken.None);

            Assert.Equal("organization-default", result.FileId);
            Assert.Equal(expectedPath, result.Path);
            Assert.True(result.Bytes > 0);
        }
        finally
        {
            Directory.Delete(templatesRoot, recursive: true);
        }
    }

    [Fact]
    public async Task MissingConfiguredDefaultTemplateFailsClosed()
    {
        var templatesRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-templates-{Guid.NewGuid():N}");
        Directory.CreateDirectory(templatesRoot);

        try
        {
            var registry = CreateRegistry(templatesRoot, "missing-template");

            var error = await Assert.ThrowsAsync<PptxValidationException>(() =>
                registry.ResolveDefaultAsync(CancellationToken.None));

            Assert.Equal("default_template_not_found", error.Code);
        }
        finally
        {
            Directory.Delete(templatesRoot, recursive: true);
        }
    }

    [Fact]
    public async Task EmptyDefaultTemplateConfigurationReturnsNull()
    {
        var templatesRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-templates-{Guid.NewGuid():N}");
        var registry = CreateRegistry(templatesRoot, string.Empty);

        var result = await registry.TryResolveDefaultAsync(CancellationToken.None);

        Assert.Null(result);
    }

    private static TemplateRegistry CreateRegistry(string templatesRoot, string defaultTemplateId)
    {
        var options = Options.Create(new PptxMcpOptions
        {
            TemplatesRoot = templatesRoot,
            DefaultTemplateId = defaultTemplateId,
            MaxFileBytes = 30 * 1024 * 1024,
            MaxSlides = 50,
        });
        return new TemplateRegistry(options, new PptxPackageGuard(options));
    }
}
