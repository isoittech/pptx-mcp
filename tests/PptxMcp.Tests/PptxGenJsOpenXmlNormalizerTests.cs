using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Compression;
using P = DocumentFormat.OpenXml.Presentation;
using PptxMcp.Presentation;

namespace PptxMcp.Tests;

public sealed class PptxGenJsOpenXmlNormalizerTests
{
    [Fact]
    public void RemovesGeneratedNotesAndRedundantPackageDirectories()
    {
        var path = TestPresentationFactory.CreateWithGeneratedNotes("VISUAL CONTENT");
        try
        {
            using (var document = PresentationDocument.Open(path, true))
            {
                var presentationPart = document.PresentationPart!;
                var notesMasterPart = presentationPart.SlideParts.Single()
                    .NotesSlidePart!
                    .NotesMasterPart!;
                presentationPart.AddPart(notesMasterPart);
                var notesMasterRelationshipId = presentationPart.GetIdOfPart(notesMasterPart);
                var notesMasterId = new P.NotesMasterId();
                notesMasterId.SetAttribute(
                    new OpenXmlAttribute(
                        "r",
                        "id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        notesMasterRelationshipId));
                presentationPart.Presentation!.InsertBefore(
                    new P.NotesMasterIdList(notesMasterId),
                    presentationPart.Presentation.GetFirstChild<P.SlideIdList>());
                presentationPart.Presentation.Save();
            }

            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                archive.CreateEntry("ppt/charts/");
                archive.CreateEntry("ppt/embeddings/");
            }

            PptxGenJsOpenXmlNormalizer.NormalizeAndValidate(path);

            using (var document = PresentationDocument.Open(path, false))
            {
                var presentationPart = document.PresentationPart!;
                Assert.Null(presentationPart.SlideParts.Single().NotesSlidePart);
                Assert.DoesNotContain(
                    presentationPart.Parts,
                    relationship => relationship.OpenXmlPart is NotesMasterPart);
                Assert.Null(
                    presentationPart.Presentation!.GetFirstChild<P.NotesMasterIdList>());
            }

            using (var archive = ZipFile.OpenRead(path))
            {
                Assert.DoesNotContain(
                    archive.Entries,
                    static entry => entry.FullName.StartsWith("ppt/notes", StringComparison.Ordinal));
                Assert.DoesNotContain(
                    archive.Entries,
                    static entry => entry.FullName.EndsWith('/'));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void MovesNotesMasterBeforeSlideList()
    {
        var slideIds = new P.SlideIdList();
        var notesMasterIds = new P.NotesMasterIdList();
        var presentation = new P.Presentation(
            new P.SlideMasterIdList(),
            slideIds,
            notesMasterIds,
            new P.SlideSize(),
            new P.NotesSize());

        var correctionCount = PptxGenJsOpenXmlNormalizer.NormalizePresentation(presentation);

        Assert.Equal(1, correctionCount);
        var children = presentation.ChildElements.ToList();
        Assert.True(
            children.IndexOf(notesMasterIds)
            < children.IndexOf(slideIds));
    }

    [Fact]
    public void NormalizesInvalidPptxGenJsTableCellMiddleAnchor()
    {
        var cellProperties = new A.TableCellProperties();
        cellProperties.SetAttribute(
            new OpenXmlAttribute(string.Empty, "anchor", string.Empty, "mid"));
        var slide = new P.Slide(cellProperties);

        var correctionCount = PptxGenJsOpenXmlNormalizer.NormalizeSlide(slide);

        Assert.Equal(1, correctionCount);
        Assert.Equal("ctr", cellProperties.GetAttribute("anchor", string.Empty).Value);
    }

    [Fact]
    public void NormalizesPptxGenJsLineAndBarChartMarkup()
    {
        var lineSeries = new C.LineChartSeries(
            new C.Index { Val = 0U },
            new C.Order { Val = 0U },
            new C.InvertIfNegative { Val = false },
            new C.DataLabels(),
            new C.Marker());
        var lineChart = new C.LineChart(
            new C.VaryColors { Val = false },
            lineSeries,
            new C.AxisId { Val = 1U },
            new C.AxisId { Val = 2U },
            new C.AxisId { Val = 3U });
        var barChart = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
            new C.BarChartSeries(
                new C.Index { Val = 0U },
                new C.Order { Val = 0U },
                new C.DataLabels(),
                new C.DataPoint(new C.Index { Val = 0U }),
                new C.DataPoint(new C.Index { Val = 1U }),
                new C.CategoryAxisData(),
                new C.Values()),
            new C.AxisId { Val = 4U },
            new C.AxisId { Val = 5U },
            new C.AxisId { Val = 6U });
        var chartSpace = new C.ChartSpace(
            new C.Chart(
                new C.PlotArea(lineChart, barChart)));

        var correctionCount = PptxGenJsOpenXmlNormalizer.NormalizeChartSpace(chartSpace);

        Assert.Equal(7, correctionCount);
        var grouping = Assert.IsType<C.Grouping>(lineChart.FirstChild);
        Assert.Equal(C.GroupingValues.Standard, grouping.Val?.Value);
        Assert.Empty(lineSeries.Elements<C.InvertIfNegative>());
        var seriesChildren = lineSeries.ChildElements.ToList();
        Assert.True(
            seriesChildren.IndexOf(lineSeries.GetFirstChild<C.Marker>()!)
            < seriesChildren.IndexOf(lineSeries.GetFirstChild<C.DataLabels>()!));
        Assert.Equal(2, lineChart.Elements<C.AxisId>().Count());
        Assert.Equal(2, barChart.Elements<C.AxisId>().Count());
        var barSeries = Assert.Single(barChart.Elements<C.BarChartSeries>());
        var barSeriesChildren = barSeries.ChildElements.ToList();
        Assert.All(
            barSeries.Elements<C.DataPoint>(),
            dataPoint => Assert.True(
                barSeriesChildren.IndexOf(dataPoint)
                < barSeriesChildren.IndexOf(barSeries.GetFirstChild<C.DataLabels>()!)));
    }

    [Fact]
    public void LeavesAlreadyNormalizedChartsUnchanged()
    {
        var lineChart = new C.LineChart(
            new C.Grouping { Val = C.GroupingValues.Standard },
            new C.VaryColors { Val = false },
            new C.AxisId { Val = 1U },
            new C.AxisId { Val = 2U });
        var chartSpace = new C.ChartSpace(
            new C.Chart(
                new C.PlotArea(lineChart)));

        var correctionCount = PptxGenJsOpenXmlNormalizer.NormalizeChartSpace(chartSpace);

        Assert.Equal(0, correctionCount);
        Assert.Equal(2, lineChart.Elements<C.AxisId>().Count());
    }
}
