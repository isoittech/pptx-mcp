using System.Text.Json;
using ModelContextProtocol.Protocol;
using PptxMcp.Domain;

namespace PptxMcp.Design;

internal static class VisualObjectPreviewResource
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);

    public static CallToolResult Create(VisualObjectPreparationView prepared)
    {
        var json = JsonSerializer.Serialize(prepared, SerializerOptions);
        return new CallToolResult
        {
            Content = [new TextContentBlock { Text = json }],
            IsError = false,
        };
    }
}
