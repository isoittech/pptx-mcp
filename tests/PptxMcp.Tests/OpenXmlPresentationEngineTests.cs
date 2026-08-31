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
    public async Task RejectsAReplacementJobThatWouldReturnAnUnchangedPresentation()
    {
        var source = TestPresentationFactory.Create("before");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var exception = await Assert.ThrowsAsync<PptxValidationException>(() =>
                new OpenXmlPresentationEngine().ReplaceTextAsync(
                    source,
                    destination,
                    [new TextReplacement("missing", "after", 1, ShapeId: 2U)],
                    CancellationToken.None));

            Assert.Equal("replacement_text_not_found", exception.Code);
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
    public async Task AnalyzeIncludesEditableTableTextWithStableGraphicFrameTarget()
    {
        var source = TestPresentationFactory.CreateWithTable(
            "Number",
            "Highest ",
            "Degree",
            "Number");
        try
        {
            var summary = await new OpenXmlPresentationEngine()
                .AnalyzeAsync(source, CancellationToken.None);

            var table = Assert.Single(Assert.Single(summary.Slides).Shapes, static shape => shape.Kind == "table");
            Assert.Equal((uint)3, table.ShapeId);
            Assert.Equal("Data Table", table.ShapeName);
            Assert.Equal("Highest DegreeNumber", table.Text);
            Assert.Equal(["Highest ", "Degree", "Number"], table.ExactTexts);
        }
        finally
        {
            File.Delete(source);
        }
    }

    [Fact]
    public async Task ReplacesTextInsideEditableTableWithoutChangingAnUntargetedShape()
    {
        var source = TestPresentationFactory.CreateWithTable(
            "Number",
            "Highest ",
            "Degree",
            "Number",
            "Number");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var result = await new OpenXmlPresentationEngine().ReplaceTextAsync(
                source,
                destination,
                [
                    new TextReplacement("Highest Degree", "最終学歴", 1, ShapeId: 3U),
                    new TextReplacement("Number", "人数", 1, ShapeId: 3U),
                ],
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var slide = document.PresentationPart!.SlideParts.Single().Slide!;
            var shapeText = string.Concat(slide.Descendants<P.Shape>().SelectMany(static shape => shape.Descendants<A.Text>()).Select(static node => node.Text));
            var tableText = string.Concat(slide.Descendants<P.GraphicFrame>().SelectMany(static frame => frame.Descendants<A.Text>()).Select(static node => node.Text));
            Assert.Equal("Number", shapeText);
            Assert.Equal("最終学歴人数人数", tableText);
            Assert.Equal(3, result.ReplacementCount);
            Assert.Equal(["/ppt/slides/slide1.xml"], result.ChangedParts);
        }
        finally
        {
            File.Delete(source);
            File.Delete(destination);
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

    [Fact]
    public async Task ComposesVisualSlidesOntoBlankCorporateLayout()
    {
        var template = TestPresentationFactory.CreateBlankBrandedTemplate();
        var visual = TestPresentationFactory.Create("VISUAL CONTENT");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var engine = new OpenXmlPresentationEngine();
            var templateSummary = await engine.AnalyzeAsync(template, CancellationToken.None);
            var result = await new OpenXmlPresentationEngine().CreateBrandedVisualDeckAsync(
                template,
                visual,
                destination,
                "auto",
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var presentationPart = document.PresentationPart!;
            var slide = Assert.Single(presentationPart.SlideParts);
            var layout = slide.SlideLayoutPart!;

            Assert.Equal(1, result.SlideCount);
            Assert.Equal("白紙（フッター有）", result.TemplateLayoutName);
            Assert.Equal(layout.Uri.ToString(), result.TemplateLayoutId);
            Assert.Equal(
                "VISUAL CONTENT",
                string.Concat(slide.Slide!.Descendants<A.Text>().Select(static text => text.Text)));
            Assert.Contains(
                "BRAND FOOTER",
                string.Concat(layout.SlideLayout!.Descendants<A.Text>().Select(static text => text.Text)),
                StringComparison.Ordinal);
            Assert.Single(presentationPart.SlideMasterParts);
            Assert.Same(presentationPart.SlideMasterParts.Single(), layout.SlideMasterPart);

            var summary = await engine.AnalyzeAsync(destination, CancellationToken.None);
            Assert.Equal(templateSummary.ValidationErrors.Count, summary.ValidationErrors.Count);
        }
        finally
        {
            File.Delete(template);
            File.Delete(visual);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task UsesConfiguredSampleSlideLayoutsForCoverAndBody()
    {
        var template = TestPresentationFactory.CreateBrandedTemplateWithSampleLayouts();
        var visual = TestPresentationFactory.CreateSlides("COVER CONTENT", "BODY CONTENT");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            await new OpenXmlPresentationEngine().CreateBrandedVisualDeckAsync(
                template,
                visual,
                destination,
                "auto",
                CancellationToken.None,
                new TemplateLayoutRolePolicy(2, 4));

            using var document = PresentationDocument.Open(destination, false);
            var presentationPart = document.PresentationPart!;
            var slides = presentationPart.Presentation!.SlideIdList!
                .Elements<P.SlideId>()
                .Select(slideId => (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!))
                .ToArray();

            Assert.Equal(2, slides.Length);
            Assert.Equal("Sample Layout 2", slides[0].SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
            Assert.Equal("Sample Layout 4", slides[1].SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
            Assert.Contains(
                "SAMPLE 2 CHROME",
                slides[0].SlideLayoutPart!.SlideLayout!.Descendants<A.Text>().Select(static text => text.Text));
            Assert.Contains(
                "SAMPLE 4 CHROME",
                slides[1].SlideLayoutPart!.SlideLayout!.Descendants<A.Text>().Select(static text => text.Text));
        }
        finally
        {
            File.Delete(template);
            File.Delete(visual);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task RejectsPlaceholderLayoutForBrandedVisualComposition()
    {
        var template = TestPresentationFactory.Create("template guide");
        var visual = TestPresentationFactory.Create("VISUAL CONTENT");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            var engine = new OpenXmlPresentationEngine();
            var layout = Assert.Single((await engine.AnalyzeAsync(template, CancellationToken.None)).Layouts);

            var error = await Assert.ThrowsAsync<PptxValidationException>(() =>
                engine.CreateBrandedVisualDeckAsync(
                    template,
                    visual,
                    destination,
                    layout.LayoutId,
                    CancellationToken.None));

            Assert.Equal("hybrid_layout_not_blank", error.Code);
        }
        finally
        {
            File.Delete(template);
            File.Delete(visual);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task RemovesEmptyGeneratedNotesPartsDuringBrandedVisualComposition()
    {
        var template = TestPresentationFactory.CreateBlankBrandedTemplate();
        var visual = TestPresentationFactory.CreateWithGeneratedNotes("VISUAL CONTENT");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            await new OpenXmlPresentationEngine().CreateBrandedVisualDeckAsync(
                template,
                visual,
                destination,
                "auto",
                CancellationToken.None);

            using (var document = PresentationDocument.Open(destination, false))
            {
                Assert.Null(document.PresentationPart!.SlideParts.Single().NotesSlidePart);
            }

            using var archive = ZipFile.OpenRead(destination);
            Assert.DoesNotContain(
                archive.Entries,
                static entry => entry.FullName.StartsWith("ppt/notesSlides/", StringComparison.Ordinal));
            Assert.DoesNotContain(
                archive.Entries,
                static entry => entry.FullName.StartsWith("ppt/notesMasters/", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(template);
            File.Delete(visual);
            File.Delete(destination);
        }
    }

    [Fact]
    public async Task PreservesSpeakerNotesDuringBrandedVisualComposition()
    {
        const string speakerNotes = "【このスライドの狙い】\n提案の価値を伝える。\n\n【トークスクリプト】\nここで承認事項を説明します。";
        var template = TestPresentationFactory.CreateBlankBrandedTemplate();
        var visual = TestPresentationFactory.CreateWithSpeakerNotes(speakerNotes, "VISUAL CONTENT");
        var destination = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        try
        {
            await new OpenXmlPresentationEngine().CreateBrandedVisualDeckAsync(
                template,
                visual,
                destination,
                "auto",
                CancellationToken.None);

            using var document = PresentationDocument.Open(destination, false);
            var presentationPart = document.PresentationPart!;
            var notesSlidePart = presentationPart.SlideParts.Single().NotesSlidePart;
            Assert.NotNull(notesSlidePart);
            Assert.Contains(
                speakerNotes,
                notesSlidePart!.NotesSlide!.Descendants<A.Text>().Select(static text => text.Text));
            Assert.Single(
                presentationPart.Parts
                    .Select(static relationship => relationship.OpenXmlPart)
                    .OfType<NotesMasterPart>());
        }
        finally
        {
            File.Delete(template);
            File.Delete(visual);
            File.Delete(destination);
        }
    }
}
