using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using PptxMcp.Presentation;

namespace PptxMcp.Tests;

public sealed class PptxGenJsOpenXmlNormalizerTests
{
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
            new C.AxisId { Val = 4U },
            new C.AxisId { Val = 5U },
            new C.AxisId { Val = 6U });
        var chartSpace = new C.ChartSpace(
            new C.Chart(
                new C.PlotArea(lineChart, barChart)));

        var correctionCount = PptxGenJsOpenXmlNormalizer.NormalizeChartSpace(chartSpace);

        Assert.Equal(5, correctionCount);
        var grouping = Assert.IsType<C.Grouping>(lineChart.FirstChild);
        Assert.Equal(C.GroupingValues.Standard, grouping.Val?.Value);
        Assert.Empty(lineSeries.Elements<C.InvertIfNegative>());
        var seriesChildren = lineSeries.ChildElements.ToList();
        Assert.True(
            seriesChildren.IndexOf(lineSeries.GetFirstChild<C.Marker>()!)
            < seriesChildren.IndexOf(lineSeries.GetFirstChild<C.DataLabels>()!));
        Assert.Equal(2, lineChart.Elements<C.AxisId>().Count());
        Assert.Equal(2, barChart.Elements<C.AxisId>().Count());
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
