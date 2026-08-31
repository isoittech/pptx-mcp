using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.IO.Compression;
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
            var presentationCorrectionCount = NormalizeGeneratedNotes(presentationPart);
            presentationCorrectionCount += NormalizePresentation(presentation);
            if (presentationCorrectionCount > 0)
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

        RemovePackageDirectoryEntries(presentationPath);

        using var validatedDocument = PresentationDocument.Open(presentationPath, false);
        var error = new OpenXmlValidator().Validate(validatedDocument).FirstOrDefault();
        if (error is not null)
        {
            throw new PptxValidationException(
                "openxml_validation_failed",
                $"The generated presentation is not valid Open XML: {FormatValidationError(error)}");
        }
    }

    internal static int NormalizeGeneratedNotes(PresentationPart presentationPart)
    {
        var correctionCount = 0;
        foreach (var slidePart in presentationPart.SlideParts.ToArray())
        {
            if (slidePart.NotesSlidePart is not { } notesSlidePart)
            {
                continue;
            }

            if (!HasSpeakerNotes(notesSlidePart))
            {
                slidePart.DeletePart(notesSlidePart);
                correctionCount++;
            }
        }

        var retainedNotesMasters = presentationPart.SlideParts
            .Select(static slidePart => slidePart.NotesSlidePart?.NotesMasterPart)
            .Where(static notesMasterPart => notesMasterPart is not null)
            .Cast<NotesMasterPart>()
            .Distinct()
            .ToArray();
        var notesMasterParts = presentationPart.Parts
            .Select(static relationship => relationship.OpenXmlPart)
            .OfType<NotesMasterPart>()
            .ToArray();
        foreach (var notesMasterPart in notesMasterParts)
        {
            if (!retainedNotesMasters.Contains(notesMasterPart))
            {
                presentationPart.DeletePart(notesMasterPart);
                correctionCount++;
            }
        }

        if (retainedNotesMasters.Length == 0)
        {
            var notesMasterIds = presentationPart.Presentation?.GetFirstChild<P.NotesMasterIdList>();
            if (notesMasterIds is not null)
            {
                notesMasterIds.Remove();
                correctionCount++;
            }

            return correctionCount;
        }

        if (retainedNotesMasters.Length != 1)
        {
            throw new PptxValidationException(
                "openxml_validation_failed",
                "The generated presentation must use exactly one notes master for retained speaker notes.");
        }

        correctionCount += EnsurePresentationNotesMasterRelationship(
            presentationPart,
            retainedNotesMasters[0]);
        return correctionCount;
    }

    internal static bool HasSpeakerNotes(NotesSlidePart notesSlidePart) =>
        notesSlidePart.NotesSlide?
            .Descendants<P.Shape>()
            .Where(static shape =>
            {
                var placeholder = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape;
                return placeholder is not null
                    && (placeholder.Type?.Value is null
                        || placeholder.Type.Value == P.PlaceholderValues.Body);
            })
            .SelectMany(static shape => shape.Descendants<A.Text>())
            .Any(static text => !string.IsNullOrWhiteSpace(text.Text)) == true;

    internal static int EnsurePresentationNotesMasterRelationship(
        PresentationPart presentationPart,
        NotesMasterPart notesMasterPart)
    {
        var correctionCount = 0;
        if (!presentationPart.Parts.Any(relationship => ReferenceEquals(relationship.OpenXmlPart, notesMasterPart)))
        {
            presentationPart.AddPart(notesMasterPart);
            correctionCount++;
        }

        var relationshipId = presentationPart.GetIdOfPart(notesMasterPart);
        var presentation = presentationPart.Presentation
            ?? throw new PptxValidationException(
                "invalid_pptx",
                "The PPTX does not contain a presentation root.");
        var notesMasterIds = presentation.GetFirstChild<P.NotesMasterIdList>();
        var currentRelationshipIds = notesMasterIds?
            .Elements<P.NotesMasterId>()
            .Select(static item => item.GetAttribute(
                "id",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value)
            .ToArray() ?? [];
        if (currentRelationshipIds.Length == 1
            && string.Equals(currentRelationshipIds[0], relationshipId, StringComparison.Ordinal))
        {
            return correctionCount;
        }

        notesMasterIds?.Remove();
        var notesMasterId = new P.NotesMasterId();
        notesMasterId.SetAttribute(new OpenXmlAttribute(
            "r",
            "id",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            relationshipId));
        var replacement = new P.NotesMasterIdList(notesMasterId);
        var slideIds = presentation.GetFirstChild<P.SlideIdList>();
        if (slideIds is null)
        {
            presentation.Append(replacement);
        }
        else
        {
            presentation.InsertBefore(replacement, slideIds);
        }

        correctionCount++;
        return correctionCount;
    }

    internal static int RemovePackageDirectoryEntries(string presentationPath)
    {
        using var archive = ZipFile.Open(presentationPath, ZipArchiveMode.Update);
        var directoryEntries = archive.Entries
            .Where(static entry => entry.FullName.EndsWith('/'))
            .ToArray();
        foreach (var directoryEntry in directoryEntries)
        {
            directoryEntry.Delete();
        }

        return directoryEntries.Length;
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
        foreach (var transform in slide.Descendants<A.Transform2D>())
        {
            var offset = transform.GetFirstChild<A.Offset>();
            var extents = transform.GetFirstChild<A.Extents>();
            if (offset is null || extents is null)
            {
                continue;
            }

            if (extents.Cx?.Value is < 0 and not long.MinValue)
            {
                offset.X = checked((offset.X?.Value ?? 0L) + extents.Cx.Value);
                extents.Cx = -extents.Cx.Value;
                transform.HorizontalFlip = !(transform.HorizontalFlip?.Value ?? false);
                correctionCount++;
            }

            if (extents.Cy?.Value is < 0 and not long.MinValue)
            {
                offset.Y = checked((offset.Y?.Value ?? 0L) + extents.Cy.Value);
                extents.Cy = -extents.Cy.Value;
                transform.VerticalFlip = !(transform.VerticalFlip?.Value ?? false);
                correctionCount++;
            }
        }

        foreach (var tableCellProperties in slide.Descendants<A.TableCellProperties>())
        {
            var anchor = tableCellProperties
                .GetAttributes()
                .FirstOrDefault(static attribute =>
                    attribute.LocalName == "anchor"
                    && string.IsNullOrEmpty(attribute.NamespaceUri));
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
