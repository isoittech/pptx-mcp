using System.Threading.Channels;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;

namespace PptxMcp.Jobs;

public sealed class JobChannel
{
    private readonly Channel<string> channel;

    public JobChannel(IOptions<PptxMcpOptions> options)
    {
        channel = Channel.CreateBounded<string>(new BoundedChannelOptions(options.Value.MaxQueueDepth)
        {
            FullMode = BoundedChannelFullMode.Wait,
            SingleReader = false,
            SingleWriter = false,
        });
    }

    public bool TryEnqueue(string jobId) => channel.Writer.TryWrite(jobId);

    public IAsyncEnumerable<string> ReadAllAsync(CancellationToken cancellationToken) =>
        channel.Reader.ReadAllAsync(cancellationToken);
}
