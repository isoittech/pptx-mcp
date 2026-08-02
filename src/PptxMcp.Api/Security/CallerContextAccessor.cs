using PptxMcp.Domain;

namespace PptxMcp.Security;

public sealed class CallerContextAccessor(IHttpContextAccessor httpContextAccessor)
{
    public CallerContext GetRequired()
    {
        var request = httpContextAccessor.HttpContext?.Request
            ?? throw new InvalidOperationException("The MCP request context is unavailable.");

        var userId = RequiredHeader(request, "X-LibreChat-User-ID");
        var conversationId = RequiredHeader(request, "X-LibreChat-Conversation-ID");
        var messageId = request.Headers["X-LibreChat-Message-ID"].FirstOrDefault();

        return new CallerContext(userId, conversationId, messageId);
    }

    private static string RequiredHeader(HttpRequest request, string name)
    {
        var value = request.Headers[name].FirstOrDefault();
        if (string.IsNullOrWhiteSpace(value) || value.Length > 256)
        {
            throw new InvalidOperationException($"Required trusted header '{name}' is missing or invalid.");
        }

        return value;
    }
}
