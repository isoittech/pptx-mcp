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

    [Fact]
    public async Task LatestResolvesUppercaseExtensionWithUnicodeAndSpacesInOriginalName()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            var expectedPath = MovePresentation(
                userDirectory,
                "uppercase-file",
                "Uppercase",
                "会社 概要.PPTX");

            var result = await CreateResolver(uploadsRoot).ResolveAsync(
                new CallerContext(userId, "conversation-1", null),
                "latest",
                CancellationToken.None);

            Assert.Equal("uppercase-file", result.FileId);
            Assert.Equal(expectedPath, result.Path);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task LatestIgnoresUnsupportedAttachmentsAndReportsNoPptx()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            await File.WriteAllTextAsync(
                Path.Combine(userDirectory, "text-file__notes.txt"),
                "not a presentation");

            var exception = await Assert.ThrowsAsync<PptxValidationException>(() =>
                CreateResolver(uploadsRoot).ResolveAsync(
                    new CallerContext(userId, "conversation-1", null),
                    "latest",
                    CancellationToken.None));

            Assert.Equal("file_not_found", exception.Code);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task ExplicitFileIdRejectsAmbiguousMatchesInsteadOfGuessing()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            MovePresentation(userDirectory, "duplicate-file", "First", "first.pptx");
            MovePresentation(userDirectory, "duplicate-file", "Second", "second.pptx");

            var exception = await Assert.ThrowsAsync<PptxValidationException>(() =>
                CreateResolver(uploadsRoot).ResolveAsync(
                    new CallerContext(userId, "conversation-1", null),
                    "duplicate-file",
                    CancellationToken.None));

            Assert.Equal("ambiguous_file_id", exception.Code);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task UploadFromAnotherUserScopeIsNotVisible()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var otherUserDirectory = Path.Combine(uploadsRoot, "user-456");
        Directory.CreateDirectory(otherUserDirectory);

        try
        {
            MovePresentation(otherUserDirectory, "private-file", "Private");

            var exception = await Assert.ThrowsAsync<PptxValidationException>(() =>
                CreateResolver(uploadsRoot).ResolveAsync(
                    new CallerContext("user-123", "conversation-1", null),
                    "private-file",
                    CancellationToken.None));

            Assert.Equal("file_not_found", exception.Code);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task LatestOnlyUsesPptxIdsAttachedToTheCurrentMessage()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            var attachedPath = MovePresentation(userDirectory, "attached-file", "Attached");
            var stalePath = MovePresentation(userDirectory, "stale-file", "Stale");
            File.SetLastWriteTimeUtc(attachedPath, DateTime.UtcNow.AddMinutes(-2));
            File.SetLastWriteTimeUtc(stalePath, DateTime.UtcNow.AddMinutes(-1));

            var result = await CreateResolver(uploadsRoot).ResolveAsync(
                new CallerContext(
                    userId,
                    "conversation-1",
                    null,
                    new HashSet<string>(["attached-file"], StringComparer.Ordinal)),
                "latest",
                CancellationToken.None);

            Assert.Equal("attached-file", result.FileId);
            Assert.Equal(attachedPath, result.Path);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task NoCurrentMessageAttachmentsDoesNotReuseAnOlderPptx()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            MovePresentation(userDirectory, "older-file", "Older");

            var exception = await Assert.ThrowsAsync<PptxValidationException>(() =>
                CreateResolver(uploadsRoot).ResolveAsync(
                    new CallerContext(
                        userId,
                        "conversation-2",
                        null,
                        new HashSet<string>(StringComparer.Ordinal)),
                    "latest",
                    CancellationToken.None));

            Assert.Equal("file_not_found", exception.Code);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    [Fact]
    public async Task ExplicitFileIdOutsideCurrentMessageAttachmentsIsNotVisible()
    {
        var uploadsRoot = Path.Combine(Path.GetTempPath(), $"pptx-mcp-uploads-{Guid.NewGuid():N}");
        var userId = "user-123";
        var userDirectory = Path.Combine(uploadsRoot, userId);
        Directory.CreateDirectory(userDirectory);

        try
        {
            MovePresentation(userDirectory, "older-file", "Older");

            var exception = await Assert.ThrowsAsync<PptxValidationException>(() =>
                CreateResolver(uploadsRoot).ResolveAsync(
                    new CallerContext(
                        userId,
                        "conversation-2",
                        null,
                        new HashSet<string>(["different-file"], StringComparer.Ordinal)),
                    "older-file",
                    CancellationToken.None));

            Assert.Equal("file_not_found", exception.Code);
        }
        finally
        {
            Directory.Delete(uploadsRoot, recursive: true);
        }
    }

    private static string MovePresentation(
        string userDirectory,
        string fileId,
        string text,
        string originalName = "template.pptx")
    {
        var target = Path.Combine(userDirectory, $"{fileId}__{originalName}");
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
