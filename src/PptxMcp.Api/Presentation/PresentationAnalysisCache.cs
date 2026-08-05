using System.Collections.Concurrent;
using System.Security.Cryptography;
using PptxMcp.Domain;

namespace PptxMcp.Presentation;

public sealed class PresentationAnalysisCache(IPresentationEngine presentationEngine)
{
    private const int MaximumEntries = 64;
    private readonly ConcurrentDictionary<string, Lazy<Task<PresentationSummary>>> cache =
        new(StringComparer.Ordinal);
    private readonly ConcurrentQueue<string> insertionOrder = new();

    public async Task<PresentationSummary> GetAsync(string sourcePath, CancellationToken cancellationToken)
    {
        var hash = await ComputeHashAsync(sourcePath, cancellationToken).ConfigureAwait(false);
        var candidate = new Lazy<Task<PresentationSummary>>(
            () => presentationEngine.AnalyzeAsync(sourcePath, CancellationToken.None),
            LazyThreadSafetyMode.ExecutionAndPublication);
        var analysis = cache.GetOrAdd(hash, candidate);

        try
        {
            var result = await analysis.Value.WaitAsync(cancellationToken).ConfigureAwait(false);
            if (ReferenceEquals(candidate, analysis))
            {
                insertionOrder.Enqueue(hash);
                Trim();
            }

            return result;
        }
        catch
        {
            if (analysis.IsValueCreated && analysis.Value.IsFaulted)
            {
                cache.TryRemove(new KeyValuePair<string, Lazy<Task<PresentationSummary>>>(hash, analysis));
            }

            throw;
        }
    }

    private void Trim()
    {
        while (cache.Count > MaximumEntries && insertionOrder.TryDequeue(out var oldestHash))
        {
            cache.TryRemove(oldestHash, out _);
        }
    }

    private static async Task<string> ComputeHashAsync(string sourcePath, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(
            sourcePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 64 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        var hash = await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(false);
        return Convert.ToHexString(hash);
    }
}
