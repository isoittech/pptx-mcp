using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class InputFileResolverTests
{
    [Fact]
    public async Task LatestResolvesMostRecentlyUploadedPptxWithinUserScope()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            var oldPath = MovePresentation(userDirectory, "old-file", "Old");
            var newPath = MovePresentation(userDirectory, "new-file", "New");
            File.SetLastWriteTimeUtc(oldPath, DateTime.UtcNow.AddMinutes(-2));
            File.SetLastWriteTimeUtc(newPath, DateTime.UtcNow.AddMinutes(-1));

            var result = await CreateResolver(uploadsRoot).ResolveAsync(
                new CallerContext(userId, "conversation-1", null),
                "latest",
                CancellationToken.None);

            Assert.Equal("new-file", result.FileId);
            Assert.Equal(newPath, result.Path);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task ExplicitFileIdStillResolvesRequestedUpload()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            var expectedPath = MovePresentation(userDirectory, "requested-file", "Requested");
            MovePresentation(userDirectory, "newer-file", "Newer");

            var result = await CreateResolver(uploadsRoot).ResolveAsync(
                new CallerContext(userId, "conversation-1", null),
                "requested-file",
                CancellationToken.None);

            Assert.Equal("requested-file", result.FileId);
            Assert.Equal(expectedPath, result.Path);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    private static string MovePresentation(string userDirectory, string fileId, string text)
    {
        var target = Path.Combine(userDirectory, $"{fileId}__template.pptx");
        File.Move(TestPresentationFactory.Create(text), target);
        return target;
    }

    private static InputFileResolver CreateResolver(string uploadsRoot)
    {
        var options = Options.Create(new PptxMcpOptions
        {
            LibreChatUploadsRoot = uploadsRoot,
            MaxFileBytes = 30 * 1024 * 1024,
            MaxSlides = 50,
        });
        return new InputFileResolver(options, new PptxPackageGuard(options));
    }
}
