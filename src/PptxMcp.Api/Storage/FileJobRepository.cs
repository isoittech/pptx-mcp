using System.Collections.Concurrent;
using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;

namespace PptxMcp.Storage;

public sealed class FileJobRepository
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
    };

    private readonly string jobsRoot;
    private readonly ConcurrentDictionary<string, SemaphoreSlim> locks = new(StringComparer.Ordinal);

    public FileJobRepository(IOptions<PptxMcpOptions> options)
    {
        jobsRoot = Path.Combine(options.Value.StorageRoot, "jobs");
        Directory.CreateDirectory(jobsRoot);
    }

    public string GetJobDirectory(string jobId) => Path.Combine(jobsRoot, ValidateJobId(jobId));

    public async Task CreateAsync(JobRecord job, CancellationToken cancellationToken)
    {
        var directory = GetJobDirectory(job.Id);
        Directory.CreateDirectory(directory);
        await WriteAsync(job, cancellationToken).ConfigureAwait(false);
    }

    public async Task<JobRecord?> GetAsync(string jobId, CancellationToken cancellationToken)
    {
        var path = Path.Combine(GetJobDirectory(jobId), "job.json");
        if (!File.Exists(path))
        {
            return null;
        }

        await using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return await JsonSerializer.DeserializeAsync<JobRecord>(stream, SerializerOptions, cancellationToken)
            .ConfigureAwait(false);
    }

    public async Task<JobRecord> UpdateAsync(
        string jobId,
        Func<JobRecord, JobRecord> update,
        CancellationToken cancellationToken)
    {
        var gate = locks.GetOrAdd(jobId, static _ => new SemaphoreSlim(1, 1));
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var current = await GetAsync(jobId, cancellationToken).ConfigureAwait(false)
                ?? throw new InvalidOperationException($"Job '{jobId}' does not exist.");
            var updated = update(current);
            await WriteAsync(updated, cancellationToken).ConfigureAwait(false);
            return updated;
        }
        finally
        {
            gate.Release();
        }
    }

    public async IAsyncEnumerable<JobRecord> ListAsync(
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (!Directory.Exists(jobsRoot))
        {
            yield break;
        }

        foreach (var directory in Directory.EnumerateDirectories(jobsRoot))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var id = Path.GetFileName(directory);
            JobRecord? job;
            try
            {
                job = await GetAsync(id, cancellationToken).ConfigureAwait(false);
            }
            catch (IOException)
            {
                continue;
            }

            if (job is not null)
            {
                yield return job;
            }
        }
    }

    public void DeleteFiles(string jobId)
    {
        var directory = GetJobDirectory(jobId);
        if (Directory.Exists(directory))
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private async Task WriteAsync(JobRecord job, CancellationToken cancellationToken)
    {
        var directory = GetJobDirectory(job.Id);
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "job.json");
        var temporaryPath = Path.Combine(directory, $"job-{Guid.NewGuid():N}.tmp");

        await using (var stream = File.Open(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            await JsonSerializer.SerializeAsync(stream, job, SerializerOptions, cancellationToken).ConfigureAwait(false);
        }

        File.Move(temporaryPath, path, overwrite: true);
    }

    private static string ValidateJobId(string jobId)
    {
        if (jobId.Length != 32 || jobId.Any(character => !Uri.IsHexDigit(character)))
        {
            throw new ArgumentException("The job identifier is invalid.", nameof(jobId));
        }

        return jobId.ToLowerInvariant();
    }
}
