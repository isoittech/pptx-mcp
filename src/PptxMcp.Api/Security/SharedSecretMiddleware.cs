using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;

namespace PptxMcp.Security;

public sealed class SharedSecretMiddleware(RequestDelegate next, IOptions<PptxMcpOptions> options)
{
    private readonly byte[] expectedSecret = Encoding.UTF8.GetBytes(options.Value.SharedSecret);

    public async Task InvokeAsync(HttpContext context)
    {
        if (context.Request.Path.StartsWithSegments("/health", StringComparison.Ordinal)
            || context.Request.Path.StartsWithSegments("/artifacts", StringComparison.Ordinal))
        {
            await next(context).ConfigureAwait(false);
            return;
        }

        var header = context.Request.Headers.Authorization.FirstOrDefault();
        var supplied = header?.StartsWith("Bearer ", StringComparison.Ordinal) == true
            ? header["Bearer ".Length..]
            : string.Empty;
        var suppliedBytes = Encoding.UTF8.GetBytes(supplied);

        if (suppliedBytes.Length != expectedSecret.Length
            || !CryptographicOperations.FixedTimeEquals(suppliedBytes, expectedSecret))
        {
            context.Response.StatusCode = StatusCodes.Status401Unauthorized;
            await context.Response.WriteAsJsonAsync(new { error = "unauthorized" }).ConfigureAwait(false);
            return;
        }

        await next(context).ConfigureAwait(false);
    }
}
