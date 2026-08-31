using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using PptxMcp.Storage;

namespace PptxMcp.Presentation;

internal static class OpenXmlRenderedDeckComposer
{
    // dom-to-pptx represents 16:9 as 13.333 inches while PowerPoint stores the
    // same nominal layout as 13 1/3 inches. Their widths differ by only 305 EMU
    // (about 0.024 pt), so treat sub-point rounding as the same slide size.
    private const int SlideSizeToleranceEmus = 12_700;

    public static async Task ComposeAsync(
        IReadOnlyList<string> sourcePaths,
        string destinationPath,
        CancellationToken cancellationToken)
    {
        if (sourcePaths.Count == 0)
        {
            throw new ArgumentException("At least one rendered deck is required.", nameof(sourcePaths));
        }

        await using (var source = File.OpenRead(sourcePaths[0]))
        await using (var destination = new FileStream(
            destinationPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            81920,
            FileOptions.Asynchronous))
        {
            await source.CopyToAsync(destination, cancellationToken).ConfigureAwait(false);
        }

        using var destinationDocument = PresentationDocument.Open(destinationPath, true);
        var destinationPresentationPart = GetPresentationPart(destinationDocument);
        var destinationPresentation = destinationPresentationPart.Presentation
            ?? throw new PptxValidationException("invalid_pptx", "The rendered presentation root is missing.");
        var destinationSlides = GetSlides(destinationPresentationPart).ToArray();
        var sharedLayout = destinationSlides.FirstOrDefault()?.SlideLayoutPart
            ?? throw new PptxValidationException("invalid_pptx", "The rendered presentation has no slide layout.");
        var sharedNotesMaster = destinationPresentationPart.Parts
            .Select(static relationship => relationship.OpenXmlPart)
            .OfType<NotesMasterPart>()
            .FirstOrDefault();
        var slideIdList = destinationPresentation.SlideIdList
            ?? destinationPresentation.AppendChild(new P.SlideIdList());
        var oldSlideIds = slideIdList.Elements<P.SlideId>().ToArray();
        var nextSlideId = Math.Max(
            256U,
            oldSlideIds.Select(static slideId => slideId.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1U);
        foreach (var slideId in oldSlideIds)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var relationshipId = slideId.RelationshipId?.Value;
            slideId.Remove();
            if (relationshipId is not null
                && destinationPresentationPart.TryGetPartById(relationshipId, out var oldSlidePart))
            {
                destinationPresentationPart.DeletePart(oldSlidePart);
            }
        }

        foreach (var sourcePath in sourcePaths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var sourceDocument = PresentationDocument.Open(sourcePath, false);
            var sourcePresentationPart = GetPresentationPart(sourceDocument);
            EnsureCompatibleSlideSize(sourcePresentationPart, destinationPresentationPart);
            var sourceSlides = GetSlides(sourcePresentationPart).ToArray();
            if (sourceSlides.Length != 1)
            {
                throw new PptxValidationException(
                    "invalid_visual_deck",
                    "Every intermediate rendered deck must contain exactly one slide.");
            }

            var sourceSlide = sourceSlides[0];
            var speakerNotes = sourceSlide.NotesSlidePart is { } notesPart
                && PptxGenJsOpenXmlNormalizer.HasSpeakerNotes(notesPart)
                && notesPart.NotesSlide is { } notes
                ? (P.NotesSlide)notes.CloneNode(deep: true)
                : null;
            if (speakerNotes is not null && sharedNotesMaster is null)
            {
                var sourceNotesMaster = sourcePresentationPart.Parts
                    .Select(static relationship => relationship.OpenXmlPart)
                    .OfType<NotesMasterPart>()
                    .SingleOrDefault()
                    ?? throw new PptxValidationException(
                        "openxml_validation_failed",
                        "Speaker notes are present but the rendered notes master is missing.");
                sharedNotesMaster = destinationPresentationPart.AddPart(sourceNotesMaster);
                PptxGenJsOpenXmlNormalizer.EnsurePresentationNotesMasterRelationship(
                    destinationPresentationPart,
                    sharedNotesMaster);
            }

            var importedSlide = destinationPresentationPart.AddPart(sourceSlide);
            if (importedSlide.NotesSlidePart is { } generatedNotesSlide)
            {
                importedSlide.DeletePart(generatedNotesSlide);
            }

            var importedLayout = importedSlide.SlideLayoutPart;
            if (importedLayout is not null && !ReferenceEquals(importedLayout, sharedLayout))
            {
                importedSlide.DeletePart(importedLayout);
            }

            if (!ReferenceEquals(importedSlide.SlideLayoutPart, sharedLayout))
            {
                importedSlide.AddPart(sharedLayout);
            }

            if (speakerNotes is not null)
            {
                AttachSpeakerNotes(importedSlide, sharedNotesMaster!, speakerNotes);
            }

            var relationshipId = destinationPresentationPart.GetIdOfPart(importedSlide);
            slideIdList.Append(new P.SlideId { Id = nextSlideId++, RelationshipId = relationshipId });
        }

        destinationPresentation.Save();
    }

    private static void AttachSpeakerNotes(
        SlidePart slidePart,
        NotesMasterPart notesMasterPart,
        P.NotesSlide notes)
    {
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        notesSlidePart.NotesSlide = notes;
        notesSlidePart.AddPart(notesMasterPart);
        notesSlidePart.AddPart(slidePart);
        notesSlidePart.NotesSlide.Save();
    }

    private static void EnsureCompatibleSlideSize(
        PresentationPart source,
        PresentationPart destination)
    {
        var sourceSize = source.Presentation?.SlideSize;
        var destinationSize = destination.Presentation?.SlideSize;
        if (sourceSize?.Cx?.Value is not int sourceWidth
            || sourceSize.Cy?.Value is not int sourceHeight
            || destinationSize?.Cx?.Value is not int destinationWidth
            || destinationSize.Cy?.Value is not int destinationHeight
            || Math.Abs(sourceWidth - destinationWidth) > SlideSizeToleranceEmus
            || Math.Abs(sourceHeight - destinationHeight) > SlideSizeToleranceEmus)
        {
            throw new PptxValidationException(
                "incompatible_slide_size",
                "Rendered slide segments must use the same slide size.");
        }
    }

    private static PresentationPart GetPresentationPart(PresentationDocument document) =>
        document.PresentationPart
        ?? throw new PptxValidationException("invalid_pptx", "The PPTX does not contain a presentation part.");

    private static IEnumerable<SlidePart> GetSlides(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? [];
        foreach (var slideId in slideIds)
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (relationshipId is not null)
            {
                yield return (SlidePart)presentationPart.GetPartById(relationshipId);
            }
        }
    }
}
