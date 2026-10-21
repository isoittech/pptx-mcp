using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxMcp.Domain;
using PptxMcp.Presentation;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class SlideExportTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), $"pptx-export-{Guid.NewGuid():N}");
    private readonly OpenXmlPresentationEngine engine = new();

    public SlideExportTests() => Directory.CreateDirectory(directory);

    [Fact]
    public async Task ExportsSelectedPagesInRequestedOrderWithoutChangingSource()
    {
        var source = Own(TestPresentationFactory.CreateSlides("Excluded cover", "Second", "Third"));
        var original = await File.ReadAllBytesAsync(source);
        var destination = Path.Combine(directory, "selected.pptx");
        var result = await engine.ExportSlidesAsync(source, destination, new SlideSelection([3, 2]), CancellationToken.None);
        var summary = await engine.AnalyzeAsync(destination, CancellationToken.None);

        Assert.Equal(2, result.SlideCount);
        Assert.Equal([3, 2], result.SourceSlideNumbers);
        Assert.Equal(["Third", "Second"], summary.Slides.Select(slide => Assert.Single(slide.Shapes).Text));
        Assert.Empty(summary.ValidationErrors);
        Assert.Equal(original, await File.ReadAllBytesAsync(source));
        using var zip = ZipFile.OpenRead(destination);
        Assert.DoesNotContain(zip.Entries, entry => entry.FullName == "ppt/slides/slide1.xml");
    }

    [Fact]
    public async Task RetainsSpeakerNotesAndTemplateLayoutForOnePage()
    {
        var source = Own(TestPresentationFactory.CreateWithSpeakerNotes("Keep the narration", "BODY CONTENT"));
        var destination = Path.Combine(directory, "notes.pptx");
        await engine.ExportSlidesAsync(source, destination, new SlideSelection([1]), CancellationToken.None);
        using var document = PresentationDocument.Open(destination, false);
        var slide = Assert.Single(document.PresentationPart!.SlideParts);
        Assert.Contains("Keep the narration", slide.NotesSlidePart!.NotesSlide!.Descendants<A.Text>().Select(node => node.Text));
        Assert.NotNull(slide.SlideLayoutPart!.SlideMasterPart);
        Assert.Empty((await engine.AnalyzeAsync(destination, CancellationToken.None)).ValidationErrors);
    }

    [Fact]
    public async Task RemovesLinksToExcludedSlidesAndCustomShowEntries()
    {
        var source = Own(TestPresentationFactory.CreateSlides("Remove me", "Keep me"));
        using (var document = PresentationDocument.Open(source, true))
        {
            var part = document.PresentationPart!;
            var ids = part.Presentation!.SlideIdList!.Elements<P.SlideId>().ToArray();
            var removed = (SlidePart)part.GetPartById(ids[0].RelationshipId!);
            var kept = (SlidePart)part.GetPartById(ids[1].RelationshipId!);
            kept.AddPart(removed, "rIdJump");
            kept.Slide!.Descendants<P.NonVisualDrawingProperties>().Last().Append(
                new A.HyperlinkOnClick { Id = "rIdJump", Action = "ppaction://hlinksldjump" });
            part.Presentation.Append(new P.CustomShowList(new P.CustomShow(
                new P.SlideList(new P.SlideListEntry { Id = ids[0].RelationshipId })) { Name = "Excluded", Id = 1U }));
        }
        var destination = Path.Combine(directory, "links.pptx");
        var result = await engine.ExportSlidesAsync(source, destination, new SlideSelection([2]), CancellationToken.None);
        Assert.Equal(1, result.RemovedSlideLinkCount);
        using var output = PresentationDocument.Open(destination, false);
        var slide = Assert.Single(output.PresentationPart!.SlideParts);
        Assert.Empty(slide.Slide!.Descendants<A.HyperlinkOnClick>());
        Assert.Null(output.PresentationPart.Presentation!.CustomShowList);
        using var zip = ZipFile.OpenRead(destination);
        Assert.DoesNotContain(zip.Entries, entry => entry.FullName == "ppt/slides/slide1.xml");
        Assert.Empty((await engine.AnalyzeAsync(destination, CancellationToken.None)).ValidationErrors);
    }

    [Theory]
    [InlineData()]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(1, 1)]
    [InlineData(3)]
    public async Task RejectsInvalidSelection(params int[] numbers)
    {
        var source = Own(TestPresentationFactory.CreateSlides("One", "Two"));
        var error = await Assert.ThrowsAsync<PptxValidationException>(() => engine.ExportSlidesAsync(
            source, Path.Combine(directory, "invalid.pptx"), new SlideSelection(numbers), CancellationToken.None));
        Assert.Equal("slide_selection_invalid", error.Code);
    }

    [Fact]
    public async Task ComposesBodyOnlyPageOnTheBodySampleLayout()
    {
        var template = Own(TestPresentationFactory.CreateBrandedTemplateWithSampleLayouts());
        var visual = Own(TestPresentationFactory.Create("ONLY BODY"));
        var destination = Path.Combine(directory, "body.pptx");
        await engine.CreateBrandedVisualDeckAsync(template, visual, destination, "auto", CancellationToken.None,
            new TemplateLayoutRolePolicy(2, 4, FirstSlideIsCover: false));
        using var document = PresentationDocument.Open(destination, false);
        var slide = Assert.Single(document.PresentationPart!.SlideParts);
        Assert.Equal("Sample Layout 4", slide.SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
        Assert.Empty((await engine.AnalyzeAsync(destination, CancellationToken.None)).ValidationErrors);
    }

    private string Own(string path)
    {
        var owned = Path.Combine(directory, Path.GetFileName(path));
        File.Move(path, owned);
        return owned;
    }

    public void Dispose() => Directory.Delete(directory, recursive: true);
}
