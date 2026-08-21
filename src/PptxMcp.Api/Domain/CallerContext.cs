using System.Security.Cryptography;
using System.Text;

namespace PptxMcp.Domain;

public sealed record CallerContext(
    string UserId,
    string ConversationId,
    string? MessageId,
    IReadOnlySet<string>? AttachmentFileIds = null)
{
    public string UserScope => Hash(UserId);

    public string ConversationScope => Hash(ConversationId);

    private static string Hash(string value)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(value));
        return Convert.ToHexString(bytes)[..24].ToLowerInvariant();
    }
}
