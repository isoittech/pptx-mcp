using System.Text.Json;
using PptxMcp.Domain;
using PptxMcp.Jobs;

namespace PptxMcp.Tests;

public sealed class ReplaceTextBatchTests
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);

    [Fact]
    public void NewPayloadPreservesIntermediateBatchState()
    {
        var serialized = JsonSerializer.SerializeToElement(new
        {
            replacements = new[]
            {
                new { find = "Company", replace = "会社", slide_number = 1, shape_id = 2 },
            },
            isFinalBatch = false,
        }, SerializerOptions);

        var payload = JobWorker.DeserializeReplaceTextPayload(serialized);

        Assert.False(payload.IsFinalBatch);
        Assert.Equal("会社", Assert.Single(payload.Replacements).Replace);
    }

    [Fact]
    public void LegacyArrayPayloadRemainsACompleteBatch()
    {
        var serialized = JsonSerializer.SerializeToElement(new[]
        {
            new TextReplacement("Company", "会社", SlideNumber: 1, ShapeId: 2),
        }, SerializerOptions);

        var payload = JobWorker.DeserializeReplaceTextPayload(serialized);

        Assert.True(payload.IsFinalBatch);
        Assert.Single(payload.Replacements);
    }
}
