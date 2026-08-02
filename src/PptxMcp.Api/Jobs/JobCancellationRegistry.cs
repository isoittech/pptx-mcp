using System.Collections.Concurrent;

namespace PptxMcp.Jobs;

public sealed class JobCancellationRegistry
{
    private readonly ConcurrentDictionary<string, CancellationTokenSource> sources = new(StringComparer.Ordinal);

    public CancellationLease Register(string jobId, TimeSpan timeout, CancellationToken hostToken)
    {
        var source = CancellationTokenSource.CreateLinkedTokenSource(hostToken);
        source.CancelAfter(timeout);
        if (!sources.TryAdd(jobId, source))
        {
            source.Dispose();
            throw new InvalidOperationException($"Job '{jobId}' is already running.");
        }

        return new CancellationLease(jobId, source, sources);
    }

    public bool Cancel(string jobId)
    {
        if (!sources.TryGetValue(jobId, out var source))
        {
            return false;
        }

        source.Cancel();
        return true;
    }

    public sealed class CancellationLease : IDisposable
    {
        private readonly string jobId;
        private readonly CancellationTokenSource source;
        private readonly ConcurrentDictionary<string, CancellationTokenSource> sources;

        internal CancellationLease(
            string jobId,
            CancellationTokenSource source,
            ConcurrentDictionary<string, CancellationTokenSource> sources)
        {
            this.jobId = jobId;
            this.source = source;
            this.sources = sources;
        }

        public CancellationToken Token => source.Token;

        public void Dispose()
        {
            sources.TryRemove(new KeyValuePair<string, CancellationTokenSource>(jobId, source));
            source.Dispose();
        }
    }
}
