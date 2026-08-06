using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using PptxMcp.Storage;

namespace PptxMcp.Presentation;

internal static class PptxGenJsOpenXmlNormalizer
{
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly HashSet<string> LineSeriesElementsAfterMarker =
    [
        "dPt",
        "dLbls",
        "trendline",
        "errBars",
        "cat",
        "val",
        "smooth",
        "extLst",
    ];
    private static readonly HashSet<string> BarSeriesElementsAfterDataPoint =
    [
        "dLbls",
        "trendline",
        "errBars",
        "cat",
        "val",
        "shape",
        "extLst",
    ];

    public static void NormalizeAndValidate(string presentationPath)
    {
        using (var document = PresentationDocument.Open(presentationPath, true))
        {
            var presentationPart = document.PresentationPart
                ?? throw new PptxValidationException(
                    "invalid_pptx",
                    "The PPTX does not contain a presentation part.");
            var presentation = presentationPart.Presentation
                ?? throw new PptxValidationException(
                    "invalid_pptx",
                    "The PPTX does not contain a presentation root.");
            if (NormalizePresentation(presentation) > 0)
            {
                presentation.Save();
            }

            foreach (var slidePart in presentationPart.SlideParts)
            {
                var slide = slidePart.Slide;
                if (slide is not null && NormalizeSlide(slide) > 0)
                {
                    slide.Save();
                }
            }

            var chartParts = presentationPart.SlideParts
                .SelectMany(static slidePart => slidePart.Parts
                    .Select(static relationship => relationship.OpenXmlPart)
                    .OfType<ChartPart>())
                .Distinct()
                .ToArray();

            foreach (var chartPart in chartParts)
            {
                var chartSpace = chartPart.ChartSpace;
                if (chartSpace is not null && NormalizeChartSpace(chartSpace) > 0)
                {
                    chartSpace.Save();
                }
            }
        }

        using var validatedDocument = PresentationDocument.Open(presentationPath, false);
        var error = new OpenXmlValidator().Validate(validatedDocument).FirstOrDefault();
        if (error is not null)
        {
            throw new PptxValidationException(
                "openxml_validation_failed",
                $"The generated presentation is not valid Open XML: {FormatValidationError(error)}");
        }
    }

    internal static int NormalizePresentation(P.Presentation presentation)
    {
        var notesMasterIds = presentation.GetFirstChild<P.NotesMasterIdList>();
        var slideIds = presentation.GetFirstChild<P.SlideIdList>();
        if (notesMasterIds is null || slideIds is null)
        {
            return 0;
        }

        var children = presentation.ChildElements.ToList();
        var notesMasterIndex = children.IndexOf(notesMasterIds);
        var slideIndex = children.IndexOf(slideIds);
        if (notesMasterIndex < slideIndex)
        {
            return 0;
        }

        notesMasterIds.Remove();
        presentation.InsertBefore(notesMasterIds, slideIds);
        return 1;
    }

    internal static int NormalizeSlide(P.Slide slide)
    {
        var correctionCount = 0;
        foreach (var tableCellProperties in slide.Descendants<A.TableCellProperties>())
        {
            var anchor = tableCellProperties.GetAttribute("anchor", string.Empty);
            if (!string.Equals(anchor.Value, "mid", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            tableCellProperties.SetAttribute(
                new OpenXmlAttribute(string.Empty, "anchor", string.Empty, "ctr"));
            correctionCount++;
        }

        return correctionCount;
    }

    internal static int NormalizeChartSpace(C.ChartSpace chartSpace)
    {
        var correctionCount = 0;
        foreach (var lineChart in chartSpace.Descendants<C.LineChart>())
        {
            if (!lineChart.Elements<C.Grouping>().Any())
            {
                lineChart.PrependChild(new C.Grouping { Val = C.GroupingValues.Standard });
                correctionCount++;
            }

            foreach (var series in lineChart.Elements<C.LineChartSeries>())
            {
                foreach (var invalidElement in series.ChildElements.Where(static child =>
                    child.LocalName == "invertIfNegative"
                    && child.NamespaceUri == ChartNamespace).ToArray())
                {
                    invalidElement.Remove();
                    correctionCount++;
                }

                correctionCount += NormalizeLineSeriesMarker(series);
            }

            correctionCount += RemoveSurplusAxisIds(lineChart.Elements<C.AxisId>());
        }

        foreach (var barChart in chartSpace.Descendants<C.BarChart>())
        {
            foreach (var series in barChart.Elements<C.BarChartSeries>())
            {
                correctionCount += NormalizeBarSeriesDataPoints(series);
            }

            correctionCount += RemoveSurplusAxisIds(barChart.Elements<C.AxisId>());
        }

        return correctionCount;
    }

    private static int NormalizeBarSeriesDataPoints(C.BarChartSeries series)
    {
        var correctionCount = 0;
        foreach (var dataPoint in series.Elements<C.DataPoint>().ToArray())
        {
            var children = series.ChildElements.ToList();
            var dataPointIndex = children.IndexOf(dataPoint);
            var firstEarlierElementThatMustFollow = children
                .Take(dataPointIndex)
                .FirstOrDefault(static child =>
                    child.NamespaceUri == ChartNamespace
                    && BarSeriesElementsAfterDataPoint.Contains(child.LocalName));
            if (firstEarlierElementThatMustFollow is null)
            {
                continue;
            }

            dataPoint.Remove();
            series.InsertBefore(dataPoint, firstEarlierElementThatMustFollow);
            correctionCount++;
        }

        return correctionCount;
    }

    private static int NormalizeLineSeriesMarker(C.LineChartSeries series)
    {
        var children = series.ChildElements.ToList();
        var marker = children.FirstOrDefault(static child =>
            child.LocalName == "marker"
            && child.NamespaceUri == ChartNamespace);
        if (marker is null)
        {
            return 0;
        }

        var markerIndex = children.IndexOf(marker);
        var firstFollowingElement = children
            .Take(markerIndex)
            .FirstOrDefault(static child =>
                child.NamespaceUri == ChartNamespace
                && LineSeriesElementsAfterMarker.Contains(child.LocalName));
        if (firstFollowingElement is null)
        {
            return 0;
        }

        marker.Remove();
        series.InsertBefore(marker, firstFollowingElement);
        return 1;
    }

    private static int RemoveSurplusAxisIds(IEnumerable<C.AxisId> axisIds)
    {
        var surplus = axisIds.Skip(2).ToArray();
        foreach (var axisId in surplus)
        {
            axisId.Remove();
        }

        return surplus.Length;
    }

    private static string FormatValidationError(ValidationErrorInfo error)
    {
        var part = error.Part?.Uri.ToString() ?? "unknown part";
        var path = error.Path?.XPath ?? "unknown path";
        return $"{part} {path}: {error.Description}";
    }
}
