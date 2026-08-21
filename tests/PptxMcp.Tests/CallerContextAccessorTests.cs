using Microsoft.AspNetCore.Http;
using PptxMcp.Security;

namespace PptxMcp.Tests;

public sealed class CallerContextAccessorTests
{
    [Fact]
    public void ReadsCurrentMessageAttachmentScopeFromTrustedHeader()
    {
        var context = CreateContext();
        context.Request.Headers["X-LibreChat-Attachment-File-IDs"] = "file-1,file_2,file-1";

        var caller = new CallerContextAccessor(new HttpContextAccessor { HttpContext = context })
            .GetRequired();

        Assert.NotNull(caller.AttachmentFileIds);
        Assert.Equal(2, caller.AttachmentFileIds.Count);
        Assert.Contains("file-1", caller.AttachmentFileIds);
        Assert.Contains("file_2", caller.AttachmentFileIds);
    }

    [Fact]
    public void NoAttachmentSentinelCreatesAnExplicitEmptyScope()
    {
        var context = CreateContext();
        context.Request.Headers["X-LibreChat-Attachment-File-IDs"] = "-";

        var caller = new CallerContextAccessor(new HttpContextAccessor { HttpContext = context })
            .GetRequired();

        Assert.NotNull(caller.AttachmentFileIds);
        Assert.Empty(caller.AttachmentFileIds);
    }

    [Fact]
    public void RejectsMalformedTrustedAttachmentScope()
    {
        var context = CreateContext();
        context.Request.Headers["X-LibreChat-Attachment-File-IDs"] = "../other-user/file";

        var accessor = new CallerContextAccessor(new HttpContextAccessor { HttpContext = context });

        Assert.Throws<InvalidOperationException>(accessor.GetRequired);
    }

    private static DefaultHttpContext CreateContext()
    {
        var context = new DefaultHttpContext();
        context.Request.Headers["X-LibreChat-User-ID"] = "user-a";
        context.Request.Headers["X-LibreChat-Conversation-ID"] = "conversation-a";
        return context;
    }
}
