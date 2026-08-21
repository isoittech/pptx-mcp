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
        var attachmentFileIds = ReadAttachmentFileIds(request);

        return new CallerContext(userId, conversationId, messageId, attachmentFileIds);
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

    private static HashSet<string>? ReadAttachmentFileIds(HttpRequest request)
    {
        var raw = request.Headers["X-LibreChat-Attachment-File-IDs"].FirstOrDefault();
        if (raw is null)
        {
            // Local clients may omit the new header. Production LibreChat always
            // supplies it, including '-' when the current request has no files.
            return null;
        }

        if (raw == "-")
        {
            return new HashSet<string>(StringComparer.Ordinal);
        }

        if (raw.Length > 4_128)
        {
            throw new InvalidOperationException("The trusted attachment file scope header is invalid.");
        }

        var fileIds = raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (fileIds.Length is < 1 or > 32
            || fileIds.Any(static fileId => !IsSafeFileId(fileId)))
        {
            throw new InvalidOperationException("The trusted attachment file scope header is invalid.");
        }

        return fileIds.ToHashSet(StringComparer.Ordinal);
    }

    private static bool IsSafeFileId(string value) =>
        value.Length is >= 1 and <= 128
        && value.All(static character => char.IsAsciiLetterOrDigit(character)
            || character is '_' or '-');
}
