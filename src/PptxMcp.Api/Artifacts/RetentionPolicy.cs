using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;

namespace PptxMcp.Artifacts;

public sealed class RetentionPolicy(IOptions<PptxMcpOptions> options)
{
    private readonly TimeSpan afterDownload = TimeSpan.FromHours(options.Value.RetentionHoursAfterDownload);

    public DateTimeOffset EffectiveExpiry(JobRecord job)
    {
        var downloadExpiry = job.FirstDownloadedAt?.Add(afterDownload);
        return downloadExpiry is null || job.ExpiresAt <= downloadExpiry
            ? job.ExpiresAt
            : downloadExpiry.Value;
    }
}
