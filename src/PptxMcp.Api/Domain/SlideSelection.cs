using System.Text.Json.Serialization;
using PptxMcp.Storage;

namespace PptxMcp.Domain;

public sealed record SlideSelection(IReadOnlyList<int> SlideNumbers)
{
    public void Validate(int maximumSlideNumber)
    {
        if (SlideNumbers is null || SlideNumbers.Count is < 1 or > 50
            || SlideNumbers.Distinct().Count() != SlideNumbers.Count
            || SlideNumbers.Any(number => number < 1 || number > maximumSlideNumber))
        {
            throw new PptxValidationException(
                "slide_selection_invalid",
                $"Provide 1 to 50 distinct slide numbers in output order, each between 1 and {maximumSlideNumber}.");
        }
    }
}

public sealed record SlideExportResult(
    [property: JsonPropertyName("slide_count")] int SlideCount,
    [property: JsonPropertyName("source_slide_numbers")] IReadOnlyList<int> SourceSlideNumbers,
    [property: JsonPropertyName("removed_slide_link_count")] int RemovedSlideLinkCount);
