using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using PptxMcp.Presentation;

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
}
