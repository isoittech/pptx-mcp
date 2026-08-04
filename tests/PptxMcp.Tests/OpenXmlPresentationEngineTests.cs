using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Compression;
using PptxMcp.Domain;
using PptxMcp.Presentation;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class OpenXmlPresentationEngineTests
{
    [Fact]
    public async Task ReplacesTextAcrossFormattingRuns()
    {
        var source = TestPresentationFactory.Create("Hel", "lo");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var engine = new OpenXmlPresentationEngine();
            var result = await engine.ReplaceTextAsync(
                source,
                destination,
                [new TextReplacement("Hello", "こんにちは", 1, "TitleBox")],
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var text = string.Concat(document.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Text>().Select(static node => node.Text));
            Assert.Equal("こんにちは", text);
            Assert.Equal(1, result.ReplacementCount);
            Assert.Single(result.ChangedParts);
        }
        finally
        {
            File.Delete(source);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task DoesNotReprocessTextIntroducedByReplacement()
    {
        var source = TestPresentationFactory.Create("a");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var result = await new OpenXmlPresentationEngine().ReplaceTextAsync(
                source,
                destination,
                [new TextReplacement("a", "aa")],
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var text = string.Concat(document.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Text>().Select(static node => node.Text));
            Assert.Equal("aa", text);
            Assert.Equal(1, result.ReplacementCount);
        }
        finally
        {
            File.Delete(source);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task RejectsOversizedReplacementInstruction()
    {
        var oversized = new string('x', 100_001);

        var exception = await Assert.ThrowsAsync<PptxValidationException>(() =>
            new OpenXmlPresentationEngine().ReplaceTextAsync(
                "not-opened.pptx",
                "not-created.pptx",
                [new TextReplacement("x", oversized)],
                CancellationToken.None));

        Assert.Equal("invalid_replacements", exception.Code);
    }

    [Fact]
    public async Task AnalyzeReturnsStableShapeTarget()
    {
        var source = TestPresentationFactory.Create("候補");
        try
        {
            var summary = await new OpenXmlPresentationEngine()
                .AnalyzeAsync(source, CancellationToken.None);

            var shape = Assert.Single(Assert.Single(summary.Slides).Shapes);
            Assert.Equal("TitleBox", shape.ShapeName);
            Assert.Equal("候補", shape.Text);
        }
        finally
        {
            File.Delete(source);
        }
    }

    [Fact]
    public async Task AnalyzeTreatsMissingPlaceholderTypeAsBody()
    {
        var source = TestPresentationFactory.Create("候補");
        try
        {
            using (var document = PresentationDocument.Open(source, true))
            {
                var layout = document.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single();
                var placeholder = layout.SlideLayout!.Descendants<P.PlaceholderShape>().Single();
                placeholder.Type = null;
                layout.SlideLayout.Save();
            }

            var summary = await new OpenXmlPresentationEngine()
                .AnalyzeAsync(source, CancellationToken.None);

            var analyzedPlaceholder = Assert.Single(Assert.Single(summary.Layouts).Placeholders);
            Assert.Equal("body", analyzedPlaceholder.PlaceholderType);
        }
        finally
        {
            File.Delete(source);
        }
    }

    [Fact]
    public async Task PopulatesPlaceholderWithoutExistingTextRun()
    {
        var source = TestPresentationFactory.Create();
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var result = await new OpenXmlPresentationEngine().PopulateTemplateAsync(
                source,
                destination,
                [new TemplateField(1, "新しい本文", ShapeName: "TitleBox")],
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var text = string.Concat(document.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Text>().Select(static node => node.Text));
            Assert.Equal("新しい本文", text);
            Assert.Equal(1, result.ReplacementCount);
        }
        finally
        {
            File.Delete(source);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task LeavesUnrelatedPackagePartsByteIdentical()
    {
        var source = TestPresentationFactory.Create("before");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            await new OpenXmlPresentationEngine().ReplaceTextAsync(
                source,
                destination,
                [new TextReplacement("before", "after")],
                CancellationToken.None);

            using var sourceArchive = ZipFile.OpenRead(source);
            using var destinationArchive = ZipFile.OpenRead(destination);
            foreach (var sourceEntry in sourceArchive.Entries.Where(static entry => entry.FullName != "ppt/slides/slide1.xml"))
            {
                var destinationEntry = destinationArchive.GetEntry(sourceEntry.FullName);
                Assert.NotNull(destinationEntry);
                await using var sourceStream = sourceEntry.Open();
                await using var destinationStream = destinationEntry.Open();
                using var sourceMemory = new MemoryStream();
                using var destinationMemory = new MemoryStream();
                await sourceStream.CopyToAsync(sourceMemory, CancellationToken.None);
                await destinationStream.CopyToAsync(destinationMemory, CancellationToken.None);
                Assert.Equal(sourceMemory.ToArray(), destinationMemory.ToArray());
            }
        }
        finally
        {
            File.Delete(source);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task CreatesDeckFromDefinedTemplateLayout()
    {
        var source = TestPresentationFactory.Create("guide");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var engine = new OpenXmlPresentationEngine();
            var layout = Assert.Single((await engine.AnalyzeAsync(source, CancellationToken.None)).Layouts);
            var result = await engine.CreateDeckAsync(
                source,
                destination,
                [
                    new DeckSlideSpec(layout.LayoutId, [new DeckField("1枚目", ShapeId: 2U)]),
                    new DeckSlideSpec(layout.LayoutId, [new DeckField("2枚目", ShapeId: 2U)]),
                ],
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var texts = document.PresentationPart!.SlideParts
                .Select(part => string.Concat(part.Slide!.Descendants<A.Text>().Select(static node => node.Text)))
                .Order(StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(2, result.SlideCount);
            Assert.Equal(["1枚目", "2枚目"], texts);
        }
        finally
        {
            File.Delete(source);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task CreatesEditableBulletAndNumberedParagraphsFromTemplateLayout()
    {
        var source = TestPresentationFactory.Create("guide");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var engine = new OpenXmlPresentationEngine();
            var layout = Assert.Single((await engine.AnalyzeAsync(source, CancellationToken.None)).Layouts);
            await engine.CreateDeckAsync(
                source,
                destination,
                [
                    new DeckSlideSpec(
                        layout.LayoutId,
                        [
                            new DeckField(
                                ShapeId: 2U,
                                Paragraphs:
                                [
                                    new DeckParagraph("優先事項", DeckParagraphKind.Bullet),
                                    new DeckParagraph("詳細項目", DeckParagraphKind.Bullet, Level: 1),
                                    new DeckParagraph("準備", DeckParagraphKind.Numbered, StartAt: 3),
                                    new DeckParagraph("実行", DeckParagraphKind.Numbered),
                                ]),
                        ]),
                ],
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var paragraphs = document.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<A.Paragraph>()
                .ToArray();

            Assert.Equal(4, paragraphs.Length);
            Assert.NotNull(paragraphs[0].ParagraphProperties?.GetFirstChild<A.CharacterBullet>());
            Assert.Equal(1, paragraphs[1].ParagraphProperties?.Level?.Value);
            var firstNumber = paragraphs[2].ParagraphProperties?.GetFirstChild<A.AutoNumberedBullet>();
            var continuedNumber = paragraphs[3].ParagraphProperties?.GetFirstChild<A.AutoNumberedBullet>();
            Assert.Equal(3, firstNumber?.StartAt?.Value);
            Assert.Null(continuedNumber?.StartAt);
            Assert.Equal(
                ["優先事項", "詳細項目", "準備", "実行"],
                paragraphs.Select(static paragraph => paragraph.Descendants<A.Text>().Single().Text).ToArray());
        }
        finally
        {
            File.Delete(source);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task RejectsAmbiguousPlainTextAndParagraphContent()
    {
        var error = await Assert.ThrowsAsync<PptxValidationException>(() =>
            new OpenXmlPresentationEngine().CreateDeckAsync(
                "not-opened.pptx",
                "not-created.pptx",
                [
                    new DeckSlideSpec(
                        "layout",
                        [
                            new DeckField(
                                "plain",
                                ShapeId: 2U,
                                Paragraphs: [new DeckParagraph("bullet", DeckParagraphKind.Bullet)]),
                        ]),
                ],
                CancellationToken.None));

        Assert.Equal("invalid_deck_spec", error.Code);
    }
}
