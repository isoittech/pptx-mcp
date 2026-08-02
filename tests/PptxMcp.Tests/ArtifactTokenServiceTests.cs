using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Security;

namespace PptxMcp.Tests;

public sealed class ArtifactTokenServiceTests
{
    [Fact]
    public void TokenIsBoundToJobFileAndExpiry()
    {
        var clock = new MutableTimeProvider(new DateTimeOffset(2026, 8, 2, 0, 0, 0, TimeSpan.Zero));
        var service = new ArtifactTokenService(
            Options.Create(new PptxMcpOptions
            {
                SigningKey = "0123456789abcdef0123456789abcdef",
                ArtifactUrlMinutes = 15,
            }),
            clock);

        var (token, _) = service.Create("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", "presentation.pptx");

        Assert.True(service.Validate("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", "presentation.pptx", token));
        Assert.False(service.Validate("bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", "presentation.pptx", token));
        Assert.False(service.Validate("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", "other.pptx", token));
        clock.Advance(TimeSpan.FromMinutes(16));
        Assert.False(service.Validate("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", "presentation.pptx", token));
    }

    [Fact]
    public void RejectsOutOfRangeExpiryWithoutThrowing()
    {
        var service = new ArtifactTokenService(
            Options.Create(new PptxMcpOptions
            {
                SigningKey = "0123456789abcdef0123456789abcdef",
            }),
            TimeProvider.System);

        Assert.False(service.Validate(
            "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
            "presentation.pptx",
            $"{long.MaxValue}.invalid"));
    }

    private sealed class MutableTimeProvider(DateTimeOffset now) : TimeProvider
    {
        private DateTimeOffset now = now;

        public override DateTimeOffset GetUtcNow() => now;

        public void Advance(TimeSpan duration) => now = now.Add(duration);
    }
}
