using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Presentation;

public sealed class OpenXmlPresentationEngine : IPresentationEngine
{
    private const int MaxAnalyzedShapes = 1_000;
    private const int MaxAnalyzedTextCharacters = 1_000;
    private const int MaxInstructionTextCharacters = 100_000;
    private const int MaxTotalInstructionTextCharacters = 2_000_000;
    private const int MaxSelectorCharacters = 512;

    public Task<PresentationSummary> AnalyzeAsync(string sourcePath, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var document = PresentationDocument.Open(sourcePath, false);
        var presentationPart = GetPresentationPart(document);
        var slides = new List<SlideSummary>();
        var hasSmartArt = false;
        var hasCharts = false;
        var hasEmbeddedWorkbook = false;
        var analyzedShapeCount = 0;
        var analysisTruncated = false;

        var slideNumber = 0;
        foreach (var slidePart in GetSlides(presentationPart))
        {
            cancellationToken.ThrowIfCancellationRequested();
            slideNumber++;
            var shapes = new List<ShapeSummary>();
            var slide = GetSlide(slidePart);

            foreach (var shape in slide.Descendants<P.Shape>())
            {
                if (analyzedShapeCount >= MaxAnalyzedShapes)
                {
                    analysisTruncated = true;
                    break;
                }

                var properties = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                var text = Truncate(JoinText(shape.Descendants<A.Text>()), ref analysisTruncated);
                shapes.Add(new ShapeSummary(
                    slideNumber,
                    properties?.Id?.Value ?? 0,
                    properties?.Name?.Value ?? string.Empty,
                    "shape",
                    string.IsNullOrEmpty(text) ? null : text));
                analyzedShapeCount++;
            }

            foreach (var frame in slide.Descendants<P.GraphicFrame>())
            {
                if (analyzedShapeCount >= MaxAnalyzedShapes)
                {
                    analysisTruncated = true;
                    break;
                }

                var properties = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                var graphicUri = frame.Graphic?.GraphicData?.Uri?.Value ?? string.Empty;
                var kind = graphicUri.EndsWith("/chart", StringComparison.Ordinal) ? "chart"
                    : graphicUri.EndsWith("/diagram", StringComparison.Ordinal) ? "smart_art"
                    : "graphic_frame";
                hasCharts |= kind == "chart";
                hasSmartArt |= kind == "smart_art";
                shapes.Add(new ShapeSummary(
                    slideNumber,
                    properties?.Id?.Value ?? 0,
                    properties?.Name?.Value ?? string.Empty,
                    kind,
                    null));
                analyzedShapeCount++;
            }

            foreach (var chartPart in slidePart.Parts.Select(static part => part.OpenXmlPart).OfType<ChartPart>())
            {
                hasCharts = true;
                hasEmbeddedWorkbook |= chartPart.Parts
                    .Select(static part => part.OpenXmlPart)
                    .OfType<EmbeddedPackagePart>()
                    .Any(IsExcelPackage);
            }

            hasEmbeddedWorkbook |= slidePart.Parts
                .Select(static part => part.OpenXmlPart)
                .OfType<EmbeddedPackagePart>()
                .Any(IsExcelPackage);

            var title = slide
                .Descendants<P.Shape>()
                .FirstOrDefault(static shape =>
                {
                    var type = shape.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .PlaceholderShape?.Type?.Value;
                    return type == P.PlaceholderValues.Title || type == P.PlaceholderValues.CenteredTitle;
                });
            var titleText = title is null ? null : Truncate(JoinText(title.Descendants<A.Text>()), ref analysisTruncated);
            var layoutName = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Name?.Value;
            slides.Add(new SlideSummary(slideNumber, titleText, layoutName, shapes));
        }

        var layouts = ReadLayouts(presentationPart);
        var theme = ReadTheme(presentationPart);
        var validationErrors = Validate(document);
        return Task.FromResult(new PresentationSummary(
            slides.Count,
            hasSmartArt,
            hasCharts,
            hasEmbeddedWorkbook,
            analysisTruncated,
            slides,
            layouts,
            theme,
            validationErrors));
    }

    public async Task<EditResult> ReplaceTextAsync(
        string sourcePath,
        string destinationPath,
        IReadOnlyList<TextReplacement> replacements,
        CancellationToken cancellationToken)
    {
        if (replacements.Count == 0
            || replacements.Count > 100
            || replacements.Any(static replacement =>
                replacement is null
                || string.IsNullOrEmpty(replacement.Find)
                || replacement.Find.Length > MaxInstructionTextCharacters
                || replacement.Replace is null
                || replacement.Replace.Length > MaxInstructionTextCharacters
                || replacement.SlideNumber is <= 0 or > 50
                || replacement.ShapeId == 0
                || replacement.ShapeName is { Length: > MaxSelectorCharacters }
                || (replacement.ShapeName is not null && string.IsNullOrWhiteSpace(replacement.ShapeName)))
            || ExceedsTotalText(replacements.SelectMany(static replacement => new[] { replacement.Find, replacement.Replace })))
        {
            throw new PptxValidationException("invalid_replacements", "Specify between 1 and 100 replacements.");
        }

        var sourceErrorCount = GetValidationErrorCount(sourcePath);
        await CopyFileAsync(sourcePath, destinationPath, cancellationToken).ConfigureAwait(false);

        var changedParts = new HashSet<string>(StringComparer.Ordinal);
        var replacementCount = 0;
        using (var document = OpenForSurgicalEdit(destinationPath))
        {
            var presentationPart = GetPresentationPart(document);
            var processedDiagramParts = new HashSet<DiagramDataPart>();
            var slideNumber = 0;
            foreach (var slidePart in GetSlides(presentationPart))
            {
                cancellationToken.ThrowIfCancellationRequested();
                slideNumber++;
                var slide = GetSlide(slidePart);
                var slideReplacementCount = 0;

                foreach (var replacement in replacements.Where(item => item.SlideNumber is null || item.SlideNumber == slideNumber))
                {
                    var matchingShapes = slide.Descendants<P.Shape>()
                        .Where(shape => replacement.ShapeName is null
                            || string.Equals(
                                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                                replacement.ShapeName,
                                StringComparison.Ordinal))
                        .Where(shape => replacement.ShapeId is null
                            || shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == replacement.ShapeId);

                    foreach (var shape in matchingShapes)
                    {
                        var count = ReplaceAcrossRuns(shape.Descendants<A.Text>().ToArray(), replacement.Find, replacement.Replace);
                        slideReplacementCount += count;
                        replacementCount += count;
                    }

                    if (replacement.ShapeName is null && replacement.ShapeId is null)
                    {
                        foreach (var diagramPart in slidePart.Parts
                            .Select(static part => part.OpenXmlPart)
                            .OfType<DiagramDataPart>()
                            .Where(processedDiagramParts.Add))
                        {
                            var root = diagramPart.RootElement;
                            if (root is not null)
                            {
                                var count = ReplaceAcrossRuns(root.Descendants<A.Text>().ToArray(), replacement.Find, replacement.Replace);
                                replacementCount += count;
                                if (count > 0)
                                {
                                    root.Save();
                                    changedParts.Add(diagramPart.Uri.ToString());
                                }
                            }
                        }
                    }
                }

                if (slideReplacementCount > 0)
                {
                    slide.Save();
                    changedParts.Add(slidePart.Uri.ToString());
                }
            }
        }

        RejectNewValidationErrors(destinationPath, sourceErrorCount);
        return new EditResult(replacementCount, changedParts.Order(StringComparer.Ordinal).ToArray());
    }

    public async Task<EditResult> PopulateTemplateAsync(
        string sourcePath,
        string destinationPath,
        IReadOnlyList<TemplateField> fields,
        CancellationToken cancellationToken)
    {
        if (fields.Count == 0 || fields.Count > 500)
        {
            throw new PptxValidationException("invalid_template_fields", "Specify between 1 and 500 template fields.");
        }

        var invalidField = fields.Any(static field =>
                field is null
                || field.SlideNumber is <= 0 or > 50
                || HasInvalidContent(field.Text, field.Paragraphs)
                || field.ShapeId == 0
                || field.ShapeName is { Length: > MaxSelectorCharacters }
                || (string.IsNullOrWhiteSpace(field.ShapeName) && field.ShapeId is null));
        var duplicate = invalidField
            ? null
            : fields.GroupBy(static field => (field.SlideNumber, field.ShapeName, field.ShapeId)).FirstOrDefault(group => group.Count() > 1);
        if (invalidField
            || duplicate is not null
            || ExceedsTotalText(fields.SelectMany(static field => GetContentText(field.Text, field.Paragraphs))))
        {
            throw new PptxValidationException("invalid_template_fields", "Template targets must be unique and valid.");
        }

        var sourceErrorCount = GetValidationErrorCount(sourcePath);
        await CopyFileAsync(sourcePath, destinationPath, cancellationToken).ConfigureAwait(false);

        var changedParts = new HashSet<string>(StringComparer.Ordinal);
        var updated = 0;
        using (var document = OpenForSurgicalEdit(destinationPath))
        {
            var slides = GetSlides(GetPresentationPart(document)).ToArray();
            foreach (var field in fields)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (field.SlideNumber > slides.Length)
                {
                    throw new PptxValidationException("template_target_not_found", $"Slide {field.SlideNumber} does not exist.");
                }

                var slidePart = slides[field.SlideNumber - 1];
                var slide = GetSlide(slidePart);
                var matchingShapes = slide.Descendants<P.Shape>()
                    .Where(candidate => field.ShapeName is null
                        || string.Equals(
                            candidate.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                            field.ShapeName,
                            StringComparison.Ordinal))
                    .Where(candidate => field.ShapeId is null
                        || candidate.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == field.ShapeId)
                    .Take(2)
                    .ToArray();
                if (matchingShapes.Length == 0)
                {
                    throw new PptxValidationException(
                        "template_target_not_found",
                        $"The requested shape was not found on slide {field.SlideNumber}.");
                }

                if (matchingShapes.Length > 1)
                {
                    throw new PptxValidationException(
                        "ambiguous_template_target",
                        $"More than one shape matched on slide {field.SlideNumber}; specify shape_id.");
                }

                SetContent(matchingShapes[0], field.Text, field.Paragraphs);
                slide.Save();
                changedParts.Add(slidePart.Uri.ToString());
                updated++;
            }
        }

        RejectNewValidationErrors(destinationPath, sourceErrorCount);
        return new EditResult(updated, changedParts.Order(StringComparer.Ordinal).ToArray());
    }

    public async Task<DeckCreationResult> CreateDeckAsync(
        string sourcePath,
        string destinationPath,
        IReadOnlyList<DeckSlideSpec> slides,
        CancellationToken cancellationToken)
    {
        if (slides.Count is <= 0 or > 50
            || slides.Any(static slide =>
                slide is null
                || string.IsNullOrWhiteSpace(slide.LayoutId)
                || slide.LayoutId.Length > MaxSelectorCharacters
                || slide.Fields is null
                || slide.Fields.Count > 100)
            || slides.Sum(static slide => slide.Fields.Count) > 1_000
            || slides.SelectMany(static slide => slide.Fields).Any(static field =>
                field is null
                || HasInvalidContent(field.Text, field.Paragraphs)
                || field.ShapeId == 0
                || field.ShapeName is { Length: > MaxSelectorCharacters }
                || (field.ShapeName is not null && string.IsNullOrWhiteSpace(field.ShapeName)))
            || ExceedsTotalText(slides
                .SelectMany(static slide => slide.Fields)
                .SelectMany(static field => GetContentText(field.Text, field.Paragraphs))))
        {
            throw new PptxValidationException("invalid_deck_spec", "Specify 1 to 50 valid slides and at most 1,000 fields.");
        }

        if (slides.SelectMany(static slide => slide.Fields).Any(static field =>
            field.ShapeId is null && field.PlaceholderIndex is null && string.IsNullOrWhiteSpace(field.ShapeName)))
        {
            throw new PptxValidationException("invalid_deck_spec", "Every field must select a placeholder by shape_id, placeholder_index, or shape_name.");
        }

        var sourceErrorCount = GetValidationErrorCount(sourcePath);
        await CopyFileAsync(sourcePath, destinationPath, cancellationToken).ConfigureAwait(false);

        var populatedFieldCount = 0;
        using (var document = OpenForSurgicalEdit(destinationPath))
        {
            var presentationPart = GetPresentationPart(document);
            var presentation = presentationPart.Presentation
                ?? throw new PptxValidationException("invalid_pptx", "The presentation root is missing.");
            var slideIdList = presentation.SlideIdList
                ?? presentation.InsertAfter(new P.SlideIdList(), presentation.HandoutMasterIdList);
            var oldSlideIds = slideIdList.Elements<P.SlideId>().ToArray();
            var nextSlideId = Math.Max(256U, oldSlideIds.Select(static slideId => slideId.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1U);
            foreach (var slideId in oldSlideIds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var relationshipId = slideId.RelationshipId?.Value;
                slideId.Remove();
                if (relationshipId is not null && presentationPart.TryGetPartById(relationshipId, out var oldSlidePart))
                {
                    presentationPart.DeletePart(oldSlidePart);
                }
            }

            var layoutParts = presentationPart.SlideMasterParts
                .SelectMany(static master => master.SlideLayoutParts)
                .ToDictionary(static layout => layout.Uri.ToString(), StringComparer.Ordinal);
            foreach (var slideSpec in slides)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!layoutParts.TryGetValue(slideSpec.LayoutId, out var layoutPart))
                {
                    throw new PptxValidationException("layout_not_found", $"Layout '{slideSpec.LayoutId}' was not found in the template.");
                }

                var layout = layoutPart.SlideLayout
                    ?? throw new PptxValidationException("invalid_template", $"Layout '{slideSpec.LayoutId}' has no layout root.");
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.AddPart(layoutPart);
                var shapeTree = CreateShapeTree(layout);
                slidePart.Slide = new P.Slide(
                    new P.CommonSlideData(shapeTree),
                    new P.ColorMapOverride(new A.MasterColorMapping()));

                foreach (var field in slideSpec.Fields)
                {
                    var matches = shapeTree.Elements<P.Shape>()
                        .Where(shape => field.ShapeName is null
                            || string.Equals(
                                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                                field.ShapeName,
                                StringComparison.Ordinal))
                        .Where(shape => field.ShapeId is null
                            || shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == field.ShapeId)
                        .Where(shape => field.PlaceholderIndex is null
                            || shape.NonVisualShapeProperties?
                                .ApplicationNonVisualDrawingProperties?
                                .PlaceholderShape?.Index?.Value == field.PlaceholderIndex)
                        .Take(2)
                        .ToArray();
                    if (matches.Length == 0)
                    {
                        throw new PptxValidationException("placeholder_not_found", $"A requested placeholder was not found in layout '{slideSpec.LayoutId}'.");
                    }

                    if (matches.Length > 1)
                    {
                        throw new PptxValidationException("ambiguous_placeholder", $"More than one placeholder matched in layout '{slideSpec.LayoutId}'; specify shape_id.");
                    }

                    SetContent(matches[0], field.Text, field.Paragraphs);
                    populatedFieldCount++;
                }

                slidePart.Slide.Save();
                var relationshipId = presentationPart.GetIdOfPart(slidePart);
                slideIdList.Append(new P.SlideId { Id = nextSlideId++, RelationshipId = relationshipId });
            }

            presentation.Save();
        }

        RejectNewValidationErrors(destinationPath, sourceErrorCount);
        return new DeckCreationResult(slides.Count, populatedFieldCount, slides.Select(static slide => slide.LayoutId).ToArray());
    }

    private static PresentationPart GetPresentationPart(PresentationDocument document) =>
        document.PresentationPart
        ?? throw new PptxValidationException("invalid_pptx", "The PPTX does not contain a presentation part.");

    private static PresentationDocument OpenForSurgicalEdit(string path) =>
        PresentationDocument.Open(path, true, new OpenSettings { AutoSave = false });

    private static IEnumerable<SlidePart> GetSlides(PresentationPart presentationPart)
    {
        var presentation = presentationPart.Presentation
            ?? throw new PptxValidationException("invalid_pptx", "The presentation root is missing.");
        var slideIds = presentation.SlideIdList?.Elements<P.SlideId>()
            ?? throw new PptxValidationException("invalid_pptx", "The PPTX does not contain a slide list.");
        foreach (var slideId in slideIds)
        {
            var relationshipId = slideId.RelationshipId?.Value
                ?? throw new PptxValidationException("invalid_pptx", "A slide relationship is missing.");
            yield return (SlidePart)presentationPart.GetPartById(relationshipId);
        }
    }

    private static P.Slide GetSlide(SlidePart slidePart) =>
        slidePart.Slide
        ?? throw new PptxValidationException("invalid_pptx", "A slide part has no slide root.");

    private static LayoutSummary[] ReadLayouts(PresentationPart presentationPart)
    {
        var layouts = new List<LayoutSummary>();
        var masterNumber = 0;
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            masterNumber++;
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                var layout = layoutPart.SlideLayout;
                if (layout is null)
                {
                    continue;
                }

                var placeholders = layout.Descendants<P.Shape>()
                    .Select(static shape => new
                    {
                        Name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                        Placeholder = shape.NonVisualShapeProperties?
                            .ApplicationNonVisualDrawingProperties?
                            .PlaceholderShape,
                    })
                    .Where(static item => item.Placeholder is not null)
                    .Select(static item => new PlaceholderSummary(
                        item.Placeholder!.Ancestors<P.Shape>().First()
                            .NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0,
                        item.Name,
                        item.Placeholder.Type?.InnerText is { Length: > 0 } rawType ? rawType : "body",
                        item.Placeholder.Index?.Value))
                    .ToArray();
                layouts.Add(new LayoutSummary(
                    layoutPart.Uri.ToString(),
                    layout.CommonSlideData?.Name?.Value ?? string.Empty,
                    masterNumber,
                    placeholders));
            }
        }

        return layouts.OrderBy(static layout => layout.MasterNumber).ThenBy(static layout => layout.LayoutId, StringComparer.Ordinal).ToArray();
    }

    private static PresentationThemeSummary? ReadTheme(PresentationPart presentationPart)
    {
        var theme = presentationPart.SlideMasterParts
            .Select(static master => master.ThemePart?.Theme)
            .FirstOrDefault(static candidate => candidate is not null);
        var themeElements = theme?.ThemeElements;
        var colorScheme = themeElements?.ColorScheme;
        if (colorScheme is null)
        {
            return null;
        }

        var colors = colorScheme.ChildElements
            .Select(static element => new ThemeColorSummary(element.LocalName, ReadThemeColor(element) ?? string.Empty))
            .Where(static color => color.Rgb.Length == 6)
            .ToArray();
        var colorsByName = colors.ToDictionary(static color => color.Name, static color => color.Rgb, StringComparer.OrdinalIgnoreCase);
        var fontScheme = themeElements?.FontScheme;
        var majorFont = fontScheme?.ChildElements.FirstOrDefault(static element => element.LocalName == "majorFont");
        var minorFont = fontScheme?.ChildElements.FirstOrDefault(static element => element.LocalName == "minorFont");
        return new PresentationThemeSummary(
            GetColor(colorsByName, "accent1"),
            GetColor(colorsByName, "accent2"),
            GetColor(colorsByName, "accent3"),
            GetColor(colorsByName, "lt1"),
            GetColor(colorsByName, "dk1"),
            ReadThemeFont(majorFont),
            ReadThemeFont(minorFont),
            colors);
    }

    private static string? ReadThemeColor(OpenXmlElement element) =>
        element.Descendants<A.RgbColorModelHex>().Select(static color => color.Val?.Value).FirstOrDefault(static value => value is not null)
        ?? element.Descendants<A.SystemColor>().Select(static color => color.LastColor?.Value).FirstOrDefault(static value => value is not null);

    private static string? ReadThemeFont(OpenXmlElement? fontCollection)
    {
        if (fontCollection is null)
        {
            return null;
        }

        var japanese = fontCollection.Descendants()
            .FirstOrDefault(static element =>
                element.LocalName == "font"
                && string.Equals(GetAttributeValue(element, "script"), "Jpan", StringComparison.OrdinalIgnoreCase));
        var japaneseTypeface = GetAttributeValue(japanese, "typeface");
        if (!string.IsNullOrWhiteSpace(japaneseTypeface))
        {
            return japaneseTypeface;
        }

        var latin = fontCollection.Descendants().FirstOrDefault(static element => element.LocalName == "latin");
        var latinTypeface = GetAttributeValue(latin, "typeface");
        return string.IsNullOrWhiteSpace(latinTypeface) ? null : latinTypeface;
    }

    private static string? GetAttributeValue(OpenXmlElement? element, string localName) =>
        element?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == localName).Value;

    private static string? GetColor(Dictionary<string, string> colors, string name) =>
        colors.TryGetValue(name, out var value) ? value : null;

    private static P.ShapeTree CreateShapeTree(P.SlideLayout layout)
    {
        var shapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));
        var layoutShapeTree = layout.CommonSlideData?.ShapeTree
            ?? throw new PptxValidationException("invalid_template", "A selected layout has no shape tree.");
        foreach (var shape in layoutShapeTree.Elements<P.Shape>().Where(static shape =>
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is not null))
        {
            shapeTree.Append(shape.CloneNode(deep: true));
        }

        return shapeTree;
    }

    private static string JoinText(IEnumerable<A.Text> textNodes) =>
        string.Concat(textNodes.Select(static node => node.Text));

    private static bool IsExcelPackage(EmbeddedPackagePart part) =>
        part.ContentType.Contains("spreadsheet", StringComparison.OrdinalIgnoreCase)
        || part.ContentType.Contains("ms-excel", StringComparison.OrdinalIgnoreCase);

    private static bool HasInvalidContent(string? text, IReadOnlyList<DeckParagraph>? paragraphs)
    {
        var hasText = text is not null;
        var hasParagraphs = paragraphs is { Count: > 0 };
        if (hasText == hasParagraphs)
        {
            return true;
        }

        if (hasText)
        {
            return text!.Length > MaxInstructionTextCharacters;
        }

        return paragraphs!.Count > 20 || paragraphs.Any(static paragraph =>
            paragraph is null
            || string.IsNullOrEmpty(paragraph.Text)
            || paragraph.Text.Length > MaxInstructionTextCharacters
            || paragraph.Level is < 0 or > 4
            || paragraph.StartAt is < 1 or > 32_767
            || (paragraph.Kind != DeckParagraphKind.Numbered && paragraph.StartAt is not null));
    }

    private static IEnumerable<string?> GetContentText(string? text, IReadOnlyList<DeckParagraph>? paragraphs)
    {
        if (text is not null)
        {
            yield return text;
        }

        if (paragraphs is not null)
        {
            foreach (var paragraph in paragraphs)
            {
                yield return paragraph.Text;
            }
        }
    }

    private static bool ExceedsTotalText(IEnumerable<string?> values)
    {
        long total = 0;
        foreach (var value in values)
        {
            total += value?.Length ?? 0;
            if (total > MaxTotalInstructionTextCharacters)
            {
                return true;
            }
        }

        return false;
    }

    private static string Truncate(string value, ref bool truncated)
    {
        if (value.Length <= MaxAnalyzedTextCharacters)
        {
            return value;
        }

        truncated = true;
        return value[..MaxAnalyzedTextCharacters];
    }

    private static int ReplaceAcrossRuns(A.Text[] nodes, string find, string replace)
    {
        if (nodes.Length == 0)
        {
            return 0;
        }

        var combined = JoinText(nodes);
        var matches = new List<int>();
        for (var searchOffset = 0; searchOffset <= combined.Length - find.Length;)
        {
            var match = combined.IndexOf(find, searchOffset, StringComparison.Ordinal);
            if (match < 0)
            {
                break;
            }

            matches.Add(match);
            searchOffset = match + find.Length;
        }

        foreach (var match in matches.AsEnumerable().Reverse())
        {
            ReplaceAt(nodes, match, find.Length, replace);
        }

        return matches.Count;
    }

    private static void ReplaceAt(A.Text[] nodes, int match, int matchLength, string replace)
    {
        var matchEnd = match + matchLength;
        var offset = 0;
        var startNode = -1;
        var endNode = -1;
        var startOffset = 0;
        var endOffset = 0;
        for (var index = 0; index < nodes.Length; index++)
        {
            var nextOffset = offset + nodes[index].Text.Length;
            if (startNode < 0 && match < nextOffset)
            {
                startNode = index;
                startOffset = match - offset;
            }

            if (matchEnd <= nextOffset)
            {
                endNode = index;
                endOffset = matchEnd - offset;
                break;
            }

            offset = nextOffset;
        }

        if (startNode < 0 || endNode < 0)
        {
            throw new InvalidOperationException("A previously located text match could not be mapped to its formatting runs.");
        }

        var prefix = nodes[startNode].Text[..startOffset];
        var suffix = nodes[endNode].Text[endOffset..];
        nodes[startNode].Text = prefix + replace + suffix;
        for (var index = startNode + 1; index <= endNode; index++)
        {
            nodes[index].Text = string.Empty;
        }
    }

    private static void SetText(P.Shape shape, string value)
    {
        var nodes = shape.Descendants<A.Text>().ToArray();
        if (nodes.Length == 0)
        {
            var textBody = shape.TextBody
                ?? throw new PptxValidationException("template_target_has_no_text", "The selected template shape has no text body.");
            var paragraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            if (paragraph is null)
            {
                paragraph = textBody.AppendChild(new A.Paragraph());
            }

            var run = new A.Run(
                new A.RunProperties { Language = "ja-JP" },
                new A.Text(value));
            paragraph.InsertBefore(run, paragraph.GetFirstChild<A.EndParagraphRunProperties>());
            return;
        }

        nodes[0].Text = value;
        foreach (var node in nodes.Skip(1))
        {
            node.Text = string.Empty;
        }
    }

    private static void SetContent(
        P.Shape shape,
        string? text,
        IReadOnlyList<DeckParagraph>? paragraphs)
    {
        if (paragraphs is { Count: > 0 })
        {
            SetParagraphs(shape, paragraphs);
            return;
        }

        SetText(shape, text ?? string.Empty);
    }

    private static void SetParagraphs(P.Shape shape, IReadOnlyList<DeckParagraph> values)
    {
        var textBody = shape.TextBody
            ?? throw new PptxValidationException("template_target_has_no_text", "The selected template shape has no text body.");
        var prototype = textBody.Elements<A.Paragraph>().FirstOrDefault();
        var prototypeProperties = prototype?.ParagraphProperties;
        var prototypeRunProperties = prototype?
            .Descendants<A.RunProperties>()
            .FirstOrDefault();
        var prototypeEndProperties = prototype?.GetFirstChild<A.EndParagraphRunProperties>();

        textBody.RemoveAllChildren<A.Paragraph>();
        foreach (var value in values)
        {
            var paragraph = new A.Paragraph();
            var paragraphProperties = prototypeProperties is null
                ? new A.ParagraphProperties()
                : (A.ParagraphProperties)prototypeProperties.CloneNode(deep: true);
            paragraphProperties.Level = value.Level;
            SetParagraphMarker(paragraphProperties, value);
            paragraph.Append(paragraphProperties);

            var runProperties = prototypeRunProperties is null
                ? new A.RunProperties { Language = "ja-JP" }
                : (A.RunProperties)prototypeRunProperties.CloneNode(deep: true);
            var text = new A.Text(value.Text);
            if (value.Text.Length > 0
                && (char.IsWhiteSpace(value.Text[0]) || char.IsWhiteSpace(value.Text[^1])))
            {
                text.SetAttribute(new OpenXmlAttribute(
                    "xml",
                    "space",
                    "http://www.w3.org/XML/1998/namespace",
                    "preserve"));
            }

            paragraph.Append(new A.Run(runProperties, text));
            if (prototypeEndProperties is not null)
            {
                paragraph.Append(prototypeEndProperties.CloneNode(deep: true));
            }

            textBody.Append(paragraph);
        }
    }

    private static void SetParagraphMarker(A.ParagraphProperties properties, DeckParagraph value)
    {
        properties.RemoveAllChildren<A.NoBullet>();
        properties.RemoveAllChildren<A.CharacterBullet>();
        properties.RemoveAllChildren<A.AutoNumberedBullet>();

        OpenXmlElement marker;
        if (value.Kind == DeckParagraphKind.Bullet)
        {
            marker = new A.CharacterBullet { Char = "•" };
        }
        else if (value.Kind == DeckParagraphKind.Numbered)
        {
            var numbered = new A.AutoNumberedBullet
            {
                Type = A.TextAutoNumberSchemeValues.ArabicPeriod,
            };
            if (value.StartAt is not null)
            {
                numbered.StartAt = value.StartAt.Value;
            }

            marker = numbered;
        }
        else
        {
            marker = new A.NoBullet();
        }
        var insertBefore = properties.ChildElements.FirstOrDefault(static child =>
            child.LocalName is "tabLst" or "defRPr" or "extLst");
        if (insertBefore is null)
        {
            properties.Append(marker);
        }
        else
        {
            properties.InsertBefore(marker, insertBefore);
        }
    }

    private static async Task CopyFileAsync(string sourcePath, string destinationPath, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("The output directory is unavailable."));
        await using var source = File.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        await using var destination = File.Open(destinationPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        await source.CopyToAsync(destination, cancellationToken).ConfigureAwait(false);
    }

    private static int GetValidationErrorCount(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        return Validate(document).Length;
    }

    private static string[] Validate(PresentationDocument document) =>
        new OpenXmlValidator()
            .Validate(document)
            .Take(100)
            .Select(static error => $"{error.Path?.XPath}: {error.Description}")
            .ToArray();

    private static void RejectNewValidationErrors(string path, int sourceErrorCount)
    {
        using var output = PresentationDocument.Open(path, false);
        var outputErrors = Validate(output);
        if (outputErrors.Length > sourceErrorCount)
        {
            throw new PptxValidationException(
                "openxml_validation_failed",
                $"The edit introduced Open XML validation errors: {outputErrors[0]}");
        }
    }
}
