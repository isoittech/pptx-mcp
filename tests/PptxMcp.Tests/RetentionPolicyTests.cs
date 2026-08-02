using Microsoft.Extensions.Options;
using PptxMcp.Artifacts;
using PptxMcp.Configuration;
using PptxMcp.Domain;

namespace PptxMcp.Tests;

public sealed class RetentionPolicyTests
{
    [Fact]
    public void FirstDownloadShortensButNeverExtendsRetention()
    {
        var policy = new RetentionPolicy(Options.Create(new PptxMcpOptions { RetentionHoursAfterDownload = 24 }));
        var created = new DateTimeOffset(2026, 8, 2, 0, 0, 0, TimeSpan.Zero);
        var job = CreateJob(created) with { FirstDownloadedAt = created.AddDays(2) };
        var lateDownload = CreateJob(created) with { FirstDownloadedAt = created.AddDays(6).AddHours(12) };

        Assert.Equal(created.AddDays(3), policy.EffectiveExpiry(job));
        Assert.Equal(created.AddDays(7), policy.EffectiveExpiry(lateDownload));
    }

    private static JobRecord CreateJob(DateTimeOffset created) => new()
    {
        Id = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
        Kind = JobKind.Analyze,
        State = JobState.Succeeded,
        UserScope = "user",
        ConversationScope = "conversation",
        SourceFileId = "file",
        CreatedAt = created,
        ExpiresAt = created.AddDays(7),
    };
}
