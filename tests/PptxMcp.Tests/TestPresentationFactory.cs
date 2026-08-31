using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests;

internal static class TestPresentationFactory
{
    public static string Create(params string[] runs)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-mcp-{Guid.NewGuid():N}.pptx");
        using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new P.Presentation();

        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        layoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(CreateShapeTree(CreateShape([], isPlaceholder: true))) { Name = "Test Layout" },
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Text,
            Preserve = true,
        };
        layoutPart.AddPart(masterPart);
        masterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(CreateShapeTree()),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            },
            new P.SlideLayoutIdList(
                new P.SlideLayoutId
                {
                    Id = 2_147_483_649U,
                    RelationshipId = masterPart.GetIdOfPart(layoutPart),
                }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(layoutPart);
        var shape = CreateShape(runs, isPlaceholder: false);
        slidePart.Slide = new P.Slide(
            new P.CommonSlideData(CreateShapeTree(shape)),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        var masterRelationshipId = presentationPart.GetIdOfPart(masterPart);
        var relationshipId = presentationPart.GetIdOfPart(slidePart);
        presentationPart.Presentation.Append(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2_147_483_648U, RelationshipId = masterRelationshipId }),
            new P.SlideIdList(new P.SlideId { Id = 256U, RelationshipId = relationshipId }),
            new P.SlideSize { Cx = 12_192_000, Cy = 6_858_000 },
            new P.NotesSize { Cx = 6_858_000, Cy = 9_144_000 });
        presentationPart.Presentation.Save();
        return path;
    }

    public static string CreateBlankBrandedTemplate()
    {
        var path = Create("template guide");
        using var document = PresentationDocument.Open(path, true);
        var layout = document.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single();
        layout.SlideLayout!.CommonSlideData = new P.CommonSlideData(
            CreateShapeTree(CreateShape(["BRAND FOOTER"], isPlaceholder: false)))
        {
            Name = "白紙（フッター有）",
        };
        layout.SlideLayout.Type = P.SlideLayoutValues.Blank;
        layout.SlideLayout.Save();
        return path;
    }

    public static string CreateBrandedTemplateWithSampleLayouts()
    {
        var path = CreateBlankBrandedTemplate();
        using var document = PresentationDocument.Open(path, true);
        var presentationPart = document.PresentationPart!;
        var presentation = presentationPart.Presentation!;
        var masterPart = presentationPart.SlideMasterParts.Single();
        var layoutIdList = masterPart.SlideMaster!.SlideLayoutIdList!;
        var slideIdList = presentation.SlideIdList!;
        var firstLayout = masterPart.SlideLayoutParts.Single();
        firstLayout.SlideLayout!.CommonSlideData!.Name = "Sample Layout 1";
        firstLayout.SlideLayout.Save();

        for (var sampleNumber = 2; sampleNumber <= 4; sampleNumber++)
        {
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new P.SlideLayout(
                new P.CommonSlideData(CreateShapeTree(CreateShape(
                    [$"SAMPLE {sampleNumber} CHROME"],
                    isPlaceholder: false)))
                {
                    Name = $"Sample Layout {sampleNumber}",
                },
                new P.ColorMapOverride(new A.MasterColorMapping()))
            {
                Type = P.SlideLayoutValues.Blank,
                Preserve = true,
            };
            layoutPart.AddPart(masterPart);
            layoutPart.SlideLayout.Save();
            layoutIdList.Append(new P.SlideLayoutId
            {
                Id = 2_147_483_648U + (uint)sampleNumber,
                RelationshipId = masterPart.GetIdOfPart(layoutPart),
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(CreateShapeTree(CreateShape(
                    [$"TEMPLATE SAMPLE {sampleNumber}"],
                    isPlaceholder: false))),
                new P.ColorMapOverride(new A.MasterColorMapping()));
            slidePart.Slide.Save();
            slideIdList.Append(new P.SlideId
            {
                Id = 255U + (uint)sampleNumber,
                RelationshipId = presentationPart.GetIdOfPart(slidePart),
            });
        }

        masterPart.SlideMaster.Save();
        presentation.Save();
        return path;
    }

    public static string CreateSlides(params string[] slideTexts)
    {
        ArgumentOutOfRangeException.ThrowIfZero(slideTexts.Length);
        var path = Create(slideTexts[0]);
        using var document = PresentationDocument.Open(path, true);
        var presentationPart = document.PresentationPart!;
        var presentation = presentationPart.Presentation!;
        var slideIdList = presentation.SlideIdList!;
        var layoutPart = presentationPart.SlideParts.Single().SlideLayoutPart!;
        for (var index = 1; index < slideTexts.Length; index++)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(CreateShapeTree(CreateShape(
                    [slideTexts[index]],
                    isPlaceholder: false))),
                new P.ColorMapOverride(new A.MasterColorMapping()));
            slidePart.Slide.Save();
            slideIdList.Append(new P.SlideId
            {
                Id = 256U + (uint)index,
                RelationshipId = presentationPart.GetIdOfPart(slidePart),
            });
        }

        presentation.Save();
        return path;
    }

    public static string CreateWithTable(string shapeText, params string[] cellTexts)
    {
        var path = Create(shapeText);
        using var document = PresentationDocument.Open(path, true);
        var slide = document.PresentationPart!.SlideParts.Single().Slide!;
        var table = new A.Table(
            new A.TableProperties { FirstRow = true },
            new A.TableGrid(cellTexts.Select(static _ => new A.GridColumn { Width = 1_800_000L })),
            new A.TableRow(cellTexts.Select(CreateTableCell)) { Height = 500_000L });
        var frame = new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = 3U, Name = "Data Table" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = 500_000L, Y = 1_000_000L },
                new A.Extents { Cx = 5_400_000L, Cy = 500_000L }),
            new A.Graphic(
                new A.GraphicData(table)
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/table",
                }));
        slide.CommonSlideData!.ShapeTree!.Append(frame);
        slide.Save();
        return path;
    }

    public static string CreateWithGeneratedNotes(params string[] runs)
    {
        var path = Create(runs);
        using var document = PresentationDocument.Open(path, true);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        var notesMasterPart = notesSlidePart.AddNewPart<NotesMasterPart>();
        notesMasterPart.NotesMaster = new P.NotesMaster(
            new P.CommonSlideData(CreateShapeTree()),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            });
        notesSlidePart.NotesSlide = new P.NotesSlide(
            new P.CommonSlideData(CreateShapeTree()),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        notesMasterPart.NotesMaster.Save();
        notesSlidePart.NotesSlide.Save();
        return path;
    }

    public static string CreateWithSpeakerNotes(string notesText, params string[] runs)
    {
        var path = Create(runs);
        using var document = PresentationDocument.Open(path, true);
        var presentationPart = document.PresentationPart!;
        var slidePart = presentationPart.SlideParts.Single();
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        var notesMasterPart = notesSlidePart.AddNewPart<NotesMasterPart>();
        notesMasterPart.NotesMaster = new P.NotesMaster(
            new P.CommonSlideData(CreateShapeTree()),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            });
        notesSlidePart.NotesSlide = new P.NotesSlide(
            new P.CommonSlideData(CreateShapeTree(CreateShape(
                [notesText],
                isPlaceholder: true,
                placeholderType: P.PlaceholderValues.Body))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        notesSlidePart.AddPart(slidePart);
        presentationPart.AddPart(notesMasterPart);
        var relationshipId = presentationPart.GetIdOfPart(notesMasterPart);
        var notesMasterId = new P.NotesMasterId();
        notesMasterId.SetAttribute(new OpenXmlAttribute(
            "r",
            "id",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            relationshipId));
        presentationPart.Presentation!.InsertBefore(
            new P.NotesMasterIdList(notesMasterId),
            presentationPart.Presentation.GetFirstChild<P.SlideIdList>());
        notesMasterPart.NotesMaster.Save();
        notesSlidePart.NotesSlide.Save();
        presentationPart.Presentation.Save();
        return path;
    }

    private static P.Shape CreateShape(IReadOnlyList<string> runs, bool isPlaceholder) =>
        CreateShape(runs, isPlaceholder, P.PlaceholderValues.Title);

    private static P.Shape CreateShape(
        IReadOnlyList<string> runs,
        bool isPlaceholder,
        P.PlaceholderValues placeholderType)
    {
        var paragraph = new A.Paragraph();
        foreach (var runText in runs)
        {
            paragraph.Append(new A.Run(new A.RunProperties { Language = "ja-JP" }, new A.Text(runText)));
        }

        var applicationProperties = new P.ApplicationNonVisualDrawingProperties();
        if (isPlaceholder)
        {
            applicationProperties.Append(new P.PlaceholderShape { Type = placeholderType, Index = 1U });
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "TitleBox" },
                new P.NonVisualShapeDrawingProperties(),
                applicationProperties),
            new P.ShapeProperties(),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), paragraph));
    }

    private static A.TableCell CreateTableCell(string text) => new(
        new A.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(
                new A.Run(
                    new A.RunProperties { Language = "en-US" },
                    new A.Text(text)))),
        new A.TableCellProperties());

    private static P.ShapeTree CreateShapeTree(params P.Shape[] shapes)
    {
        var shapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));
        shapeTree.Append(shapes);
        return shapeTree;
    }
}
