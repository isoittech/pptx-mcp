using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;

namespace PptxMcp.Security;

public sealed class ArtifactTokenService(IOptions<PptxMcpOptions> options, TimeProvider timeProvider)
{
    private readonly byte[] signingKey = Encoding.UTF8.GetBytes(options.Value.SigningKey);
    private readonly TimeSpan lifetime = TimeSpan.FromMinutes(options.Value.ArtifactUrlMinutes);

    public (string Token, DateTimeOffset ExpiresAt) Create(string jobId, string fileName)
    {
        var expiresAt = timeProvider.GetUtcNow().Add(lifetime);
        var expires = expiresAt.ToUnixTimeSeconds().ToString(CultureInfo.InvariantCulture);
        var signature = Sign(jobId, fileName, expires);
        return ($"{expires}.{WebEncoders.Base64UrlEncode(signature)}", expiresAt);
    }

    public bool Validate(string jobId, string fileName, string token)
    {
        var separator = token.IndexOf('.', StringComparison.Ordinal);
        if (separator <= 0 || separator == token.Length - 1)
        {
            return false;
        }

        var expires = token[..separator];
        if (!long.TryParse(expires, CultureInfo.InvariantCulture, out var unixSeconds)
            || unixSeconds < DateTimeOffset.MinValue.ToUnixTimeSeconds()
            || unixSeconds > DateTimeOffset.MaxValue.ToUnixTimeSeconds()
            || DateTimeOffset.FromUnixTimeSeconds(unixSeconds) < timeProvider.GetUtcNow())
        {
            return false;
        }

        byte[] supplied;
        try
        {
            supplied = WebEncoders.Base64UrlDecode(token[(separator + 1)..]);
        }
        catch (FormatException)
        {
            return false;
        }

        var expected = Sign(jobId, fileName, expires);
        return supplied.Length == expected.Length
            && CryptographicOperations.FixedTimeEquals(supplied, expected);
    }

    private byte[] Sign(string jobId, string fileName, string expires)
    {
        var data = Encoding.UTF8.GetBytes($"{jobId}\n{fileName}\n{expires}");
        return HMACSHA256.HashData(signingKey, data);
    }
}
