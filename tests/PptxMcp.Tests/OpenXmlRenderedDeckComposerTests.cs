using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using PptxMcp.Presentation;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class OpenXmlRenderedDeckComposerTests
{
    [Fact]
    public async Task ComposesOneSlideRendererOutputsInOrderAndPreservesSpeakerNotes()
    {
        const string speakerNotes = "【このスライドの狙い】\n混在描画を確認する。\n\n【トークスクリプト】\nDOMとネイティブ図を同じ資料に残します。";
        var domSlide = TestPresentationFactory.Create("DOM CONTENT");
        var nativeSlide = TestPresentationFactory.CreateWithSpeakerNotes(speakerNotes, "NATIVE CHART CONTENT");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-composed-{Guid.NewGuid():N}.pptx");
        try
        {
            await OpenXmlRenderedDeckComposer.ComposeAsync(
                [domSlide, nativeSlide],
                destination,
                CancellationToken.None);
            PptxGenJsOpenXmlNormalizer.NormalizeAndValidate(destination);

            using var document = PresentationDocument.Open(destination, false);
            var slides = document.PresentationPart!.SlideParts.ToArray();
            Assert.Equal(2, slides.Length);
            Assert.Equal(
                "DOM CONTENT",
                string.Concat(slides[0].Slide!.Descendants<A.Text>().Select(static text => text.Text)));
            Assert.Equal(
                "NATIVE CHART CONTENT",
                string.Concat(slides[1].Slide!.Descendants<A.Text>().Select(static text => text.Text)));
            Assert.Contains(
                speakerNotes,
                slides[1].NotesSlidePart!.NotesSlide!.Descendants<A.Text>().Select(static text => text.Text));
            Assert.Single(
                document.PresentationPart.Parts
                    .Select(static relationship => relationship.OpenXmlPart)
                    .OfType<NotesMasterPart>());
        }
        finally
        {
            File.Delete(domSlide);
            File.Delete(nativeSlide);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task ComposesEquivalentWideSlideSizesWithRendererRounding()
    {
        var firstSlide = TestPresentationFactory.Create("POWERPOINT WIDE");
        var roundedDomSlide = TestPresentationFactory.Create("DOM WIDE");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-rounded-wide-{Guid.NewGuid():N}.pptx");
        try
        {
            SetSlideWidth(roundedDomSlide, 12_191_695);

            await OpenXmlRenderedDeckComposer.ComposeAsync(
                [firstSlide, roundedDomSlide],
                destination,
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            Assert.Equal(12_192_000, document.PresentationPart!.Presentation!.SlideSize!.Cx!.Value);
            Assert.Equal(2, document.PresentationPart.SlideParts.Count());
        }
        finally
        {
            File.Delete(firstSlide);
            File.Delete(roundedDomSlide);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task RejectsMeaningfullyDifferentSlideSizes()
    {
        var firstSlide = TestPresentationFactory.Create("POWERPOINT WIDE");
        var differentSlide = TestPresentationFactory.Create("DIFFERENT SIZE");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-different-size-{Guid.NewGuid():N}.pptx");
        try
        {
            SetSlideWidth(differentSlide, 12_150_000);

            var error = await Assert.ThrowsAsync<PptxValidationException>(() =>
                OpenXmlRenderedDeckComposer.ComposeAsync(
                    [firstSlide, differentSlide],
                    destination,
                    CancellationToken.None));

            Assert.Equal("incompatible_slide_size", error.Code);
        }
        finally
        {
            File.Delete(firstSlide);
            File.Delete(differentSlide);
            File.Delete(destination);
        }
    }

    private static void SetSlideWidth(string path, int width)
    {
        using var document = PresentationDocument.Open(path, true);
        document.PresentationPart!.Presentation!.SlideSize!.Cx = width;
        document.PresentationPart.Presentation.Save();
    }
}
