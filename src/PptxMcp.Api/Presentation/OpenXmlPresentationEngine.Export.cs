using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Presentation;

public sealed partial class OpenXmlPresentationEngine
{
    public async Task<SlideExportResult> ExportSlidesAsync(
        string sourcePath,
        string destinationPath,
        SlideSelection selection,
        CancellationToken cancellationToken)
    {
        selection.Validate(50);
        var sourceErrorCount = GetValidationErrorCount(sourcePath);
        await CopyFileAsync(sourcePath, destinationPath, cancellationToken).ConfigureAwait(false);
        var removedLinks = 0;
        using (var document = OpenForSurgicalEdit(destinationPath))
        {
            var presentationPart = GetPresentationPart(document);
            var presentation = presentationPart.Presentation!;
            var slides = GetSlides(presentationPart).ToArray();
            selection.Validate(slides.Length);
            var slideIds = presentation.SlideIdList!.Elements<P.SlideId>().ToArray();
            var selected = selection.SlideNumbers.ToHashSet();
            var removedParts = slides.Where((_, index) => !selected.Contains(index + 1))
                .Cast<OpenXmlPart>().ToHashSet();
            var removedIds = slideIds.Where((_, index) => !selected.Contains(index + 1))
                .Select(static id => id.Id!.Value).ToHashSet();
            var removedRelationships = slideIds.Where((_, index) => !selected.Contains(index + 1))
                .Select(static id => id.RelationshipId!.Value!).ToHashSet(StringComparer.Ordinal);

            foreach (var show in presentation.CustomShowList?.Elements<P.CustomShow>().ToArray() ?? [])
            {
                foreach (var entry in show.Descendants<P.SlideListEntry>().ToArray())
                {
                    if (removedRelationships.Contains(entry.Id?.Value ?? string.Empty))
                    {
                        entry.Remove();
                    }
                }
                if (!show.Descendants<P.SlideListEntry>().Any())
                {
                    show.Remove();
                }
            }
            if (presentation.CustomShowList is { HasChildren: false } emptyShows)
            {
                emptyShows.Remove();
            }
            foreach (var entry in presentation.Descendants<P14.SectionSlideIdListEntry>().ToArray())
            {
                if (entry.Id is not null && removedIds.Contains(entry.Id.Value))
                {
                    entry.Remove();
                }
            }
            foreach (var section in presentation.Descendants<P14.Section>().ToArray())
            {
                if (!section.Descendants<P14.SectionSlideIdListEntry>().Any())
                {
                    section.Remove();
                }
            }

            // Slide hyperlinks must not keep an excluded slide reachable inside the output package.
            foreach (var part in EnumerateParts(document))
            {
                if (part == presentationPart || removedParts.Contains(part))
                {
                    continue;
                }
                foreach (var relationship in part.Parts.Where(pair => removedParts.Contains(pair.OpenXmlPart)).ToArray())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var root = part.RootElement;
                    if (root is not null)
                    {
                        foreach (var link in root.Descendants<A.HyperlinkType>()
                            .Where(link => link.Id?.Value == relationship.RelationshipId).ToArray())
                        {
                            link.Remove();
                            removedLinks++;
                        }
                        if (root.Descendants().Any(element => element.GetAttributes().Any(attribute =>
                            attribute.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            && attribute.Value == relationship.RelationshipId)))
                        {
                            throw new PptxValidationException(
                                "slide_selection_reference_unsupported",
                                "A selected part has an unsupported reference to an excluded slide. Include that slide or remove the reference before exporting.");
                        }
                        root.Save();
                    }
                    part.DeletePart(relationship.RelationshipId);
                }
            }

            presentation.SlideIdList.RemoveAllChildren();
            foreach (var number in selection.SlideNumbers)
            {
                presentation.SlideIdList.Append(slideIds[number - 1]);
            }
            foreach (var part in removedParts)
            {
                cancellationToken.ThrowIfCancellationRequested();
                presentationPart.DeletePart(part);
            }
            // Whole-deck summaries and the original cover thumbnail may disclose excluded content.
            if (document.ExtendedFilePropertiesPart is { } properties)
            {
                document.DeletePart(properties);
            }
            if (document.ThumbnailPart is { } thumbnail)
            {
                document.DeletePart(thumbnail);
            }
            presentation.Save();
        }
        RejectNewValidationErrors(destinationPath, sourceErrorCount);
        return new SlideExportResult(selection.SlideNumbers.Count, selection.SlideNumbers.ToArray(), removedLinks);
    }

    private static OpenXmlPart[] EnumerateParts(OpenXmlPartContainer container)
    {
        var visited = new HashSet<OpenXmlPart>();
        var pending = new Queue<OpenXmlPart>(container.Parts.Select(static pair => pair.OpenXmlPart));
        while (pending.TryDequeue(out var part))
        {
            if (!visited.Add(part))
            {
                continue;
            }
            foreach (var child in part.Parts)
            {
                pending.Enqueue(child.OpenXmlPart);
            }
        }
        return visited.ToArray();
    }
}
