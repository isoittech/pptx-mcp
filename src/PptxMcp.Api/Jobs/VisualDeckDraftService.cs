using System.Collections.Concurrent;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Jobs;

public sealed class VisualDeckDraftService(
    IOptions<PptxMcpOptions> options,
    TimeProvider timeProvider)
{
    public const int MaximumBatchSlides = 4;
    private const int MaximumActiveDrafts = 128;
    private static readonly TimeSpan DraftLifetime = TimeSpan.FromHours(1);
    private readonly ConcurrentDictionary<string, DraftState> drafts = new(StringComparer.Ordinal);
    private readonly object beginLock = new();
    private readonly PptxMcpOptions options = options.Value;

    public VisualDeckDraftView Begin(
        CallerContext caller,
        string title,
        int expectedSlideCount,
        VisualThemeSpec? theme,
        string? subject,
        string language,
        VisualDesignSpec? design,
        string templateSourceFileId = "default",
        string templateLayoutId = "auto",
        bool allowSubmittedReplacement = false,
        ValidatedDesignBriefBinding? designBrief = null)
    {
        VisualDeckValidator.ValidateMetadata(title, subject, language, theme, design);
        if (expectedSlideCount is < 1 || expectedSlideCount > options.MaxSlides)
        {
            throw new PptxValidationException(
                "visual_draft_slide_count_invalid",
                $"expectedSlideCount must be between 1 and {options.MaxSlides}.");
        }

        ValidateTemplateSelection(templateSourceFileId, templateLayoutId);

        lock (beginLock)
        {
            PruneExpired();
            var active = drafts.Values
                .Where(draft =>
                    string.Equals(draft.UserScope, caller.UserScope, StringComparison.Ordinal)
                    && string.Equals(draft.ConversationScope, caller.ConversationScope, StringComparison.Ordinal))
                .OrderByDescending(static draft => draft.CreatedAt)
                .FirstOrDefault();
            if (active is not null && active.Status is DraftStatus.Collecting or DraftStatus.Submitting)
            {
                if (MatchesActiveDraft(
                        active,
                        title,
                        expectedSlideCount,
                        theme,
                        subject,
                        language,
                        design,
                        templateSourceFileId,
                        templateLayoutId,
                        designBrief))
                {
                    return CreateView(active);
                }

                throw new PptxValidationException(
                    "visual_draft_already_active",
                    "A different Visual Deck draft is already active in this conversation. Continue it with the returned draft_id; do not replace its locked title, slide count, theme, design, template, or Design Brief.");
            }

            if (active is not null && !allowSubmittedReplacement)
            {
                throw new PptxValidationException(
                    "visual_draft_already_submitted",
                    $"This conversation already submitted visual deck job {active.JobId}. Do not start the deck again.");
            }

            if (drafts.Count >= MaximumActiveDrafts)
            {
                throw new PptxValidationException(
                    "visual_draft_capacity_reached",
                    "Too many visual deck drafts are active. Retry later.");
            }

            var now = timeProvider.GetUtcNow();
            while (true)
            {
                var draft = new DraftState(
                    Guid.NewGuid().ToString("N"),
                    caller.UserScope,
                    caller.ConversationScope,
                    title,
                    expectedSlideCount,
                    [],
                    theme,
                    subject,
                    language,
                    design,
                    NormalizeTemplateSource(templateSourceFileId),
                    NormalizeTemplateLayout(templateSourceFileId, templateLayoutId),
                    designBrief,
                    DraftStatus.Collecting,
                    null,
                    now,
                    now.Add(DraftLifetime));
                if (drafts.TryAdd(draft.Id, draft))
                {
                    return CreateView(draft);
                }
            }
        }
    }

    public VisualDeckDraftView AddSlides(
        CallerContext caller,
        string draftId,
        int? startSlideNumber,
        IReadOnlyList<VisualSlideSpec> slides)
    {
        if (slides is null || slides.Count is < 1 or > MaximumBatchSlides || slides.Any(static slide => slide is null))
        {
            throw new PptxValidationException(
                "visual_draft_batch_invalid",
                $"Provide between 1 and {MaximumBatchSlides} complete slides per draft batch.");
        }

        while (true)
        {
            var current = GetOwned(caller, draftId);
            EnsureCollecting(current);
            var expectedStart = current.Slides.Count + 1;
            if (startSlideNumber.HasValue && startSlideNumber.Value != expectedStart)
            {
                throw new PptxValidationException(
                    "visual_draft_position_invalid",
                    $"startSlideNumber must be {expectedStart} for this draft. Do not resend an accepted batch.");
            }

            ValidateBrandRecipes(current, slides, expectedStart);

            var combined = current.Slides.Concat(slides).ToArray();
            if (combined.Length > current.ExpectedSlideCount)
            {
                throw new PptxValidationException(
                    "visual_draft_slide_count_exceeded",
                    $"This draft expects exactly {current.ExpectedSlideCount} slides; the batch would produce {combined.Length}.");
            }

            VisualDeckValidator.Validate(CreateDeck(current, combined), options.MaxSlides);
            var updated = current with { Slides = combined };
            if (drafts.TryUpdate(current.Id, updated, current))
            {
                return CreateView(updated);
            }
        }
    }

    public VisualDeckDraftSubmission AcquireForSubmission(
        CallerContext caller,
        string draftId,
        string? requestedTemplateSourceFileId = null,
        string? requestedTemplateLayoutId = null)
    {
        while (true)
        {
            var current = GetOwned(caller, draftId);
            EnsureTemplateSelectionMatches(
                current,
                requestedTemplateSourceFileId,
                requestedTemplateLayoutId);
            if (current.Status == DraftStatus.Submitted && current.JobId is not null)
            {
                return new VisualDeckDraftSubmission(
                    null,
                    current.JobId,
                    current.TemplateSourceFileId,
                    current.TemplateLayoutId);
            }

            if (current.Status == DraftStatus.Submitting)
            {
                throw new PptxValidationException(
                    "visual_draft_submission_in_progress",
                    "This visual deck draft is already being submitted. Do not call a finish tool again.");
            }

            if (current.Slides.Count != current.ExpectedSlideCount)
            {
                throw new PptxValidationException(
                    "visual_draft_incomplete",
                    $"The draft contains {current.Slides.Count} of {current.ExpectedSlideCount} slides. Add the remaining slides before finishing.");
            }

            var deck = CreateDeck(current, current.Slides);
            VisualDeckValidator.Validate(deck, options.MaxSlides);
            var submitting = current with { Status = DraftStatus.Submitting };
            if (drafts.TryUpdate(current.Id, submitting, current))
            {
                return new VisualDeckDraftSubmission(
                    deck,
                    null,
                    current.TemplateSourceFileId,
                    current.TemplateLayoutId);
            }
        }
    }

    public void MarkSubmitted(CallerContext caller, string draftId, string jobId)
    {
        while (true)
        {
            var current = GetOwned(caller, draftId);
            if (current.Status == DraftStatus.Submitted && string.Equals(current.JobId, jobId, StringComparison.Ordinal))
            {
                return;
            }

            if (current.Status != DraftStatus.Submitting)
            {
                throw new InvalidOperationException("The visual deck draft is not awaiting submission completion.");
            }

            var submitted = current with { Status = DraftStatus.Submitted, JobId = jobId };
            if (drafts.TryUpdate(current.Id, submitted, current))
            {
                return;
            }
        }
    }

    public void ReleaseSubmission(CallerContext caller, string draftId)
    {
        while (true)
        {
            var current = GetOwned(caller, draftId);
            if (current.Status != DraftStatus.Submitting)
            {
                return;
            }

            if (drafts.TryUpdate(current.Id, current with { Status = DraftStatus.Collecting }, current))
            {
                return;
            }
        }
    }

    private DraftState GetOwned(CallerContext caller, string draftId)
    {
        if (!IsValidId(draftId)
            || !drafts.TryGetValue(draftId.ToLowerInvariant(), out var draft)
            || !string.Equals(draft.UserScope, caller.UserScope, StringComparison.Ordinal)
            || !string.Equals(draft.ConversationScope, caller.ConversationScope, StringComparison.Ordinal))
        {
            throw new PptxValidationException(
                "visual_draft_not_found",
                "The visual deck draft was not found in this conversation. Start a new draft.");
        }

        if (draft.ExpiresAt <= timeProvider.GetUtcNow())
        {
            drafts.TryRemove(draft.Id, out _);
            throw new PptxValidationException(
                "visual_draft_expired",
                "The visual deck draft expired. Start a new draft.");
        }

        return draft;
    }

    private static void EnsureCollecting(DraftState draft)
    {
        if (draft.Status != DraftStatus.Collecting)
        {
            throw new PptxValidationException(
                "visual_draft_not_editable",
                "The visual deck draft has already been submitted or is being submitted.");
        }
    }

    private static VisualDeckSpec CreateDeck(DraftState draft, IReadOnlyList<VisualSlideSpec> slides) =>
        new(
            draft.Title,
            slides,
            draft.Theme,
            draft.Subject,
            draft.Language,
            draft.Design,
            "visual-v5",
            CreatePersistedBrandBinding(draft));

    private static VisualDeckBrandProfileBinding? CreatePersistedBrandBinding(DraftState draft)
    {
        if (draft.DesignBrief is null)
        {
            return null;
        }

        var slideBindings = draft.DesignBrief.AssetPlan
            .OrderBy(static pair => pair.Key)
            .Select(pair =>
            {
                var recipe = draft.DesignBrief.Profile.LayoutRecipes.Single(item =>
                    string.Equals(item.Id, pair.Value.RecipeId, StringComparison.Ordinal));
                return new VisualSlideRecipeBinding(
                    pair.Key,
                    recipe.Id,
                    recipe.SemanticKind,
                    recipe.Density,
                    recipe.Variant);
            })
            .ToArray();
        var assetAudit = draft.DesignBrief.AssetPlan
            .OrderBy(static pair => pair.Key)
            .Select(pair => new VisualSlideAssetAudit(
                pair.Key,
                pair.Value.VisualPurpose,
                pair.Value.PreferredMedium,
                pair.Value.Acquisition,
                pair.Value.Fallback,
                pair.Value.Status,
                pair.Value.LicenseStatus,
                pair.Value.ApprovedAssetCollectionId,
                pair.Value.AttributionRef,
                pair.Value.CropIntent,
                pair.Value.AspectRatio,
                pair.Value.TextSafeArea))
            .ToArray();
        var audit = new VisualDeckDesignBriefAudit(
            draft.DesignBrief.Brief.SourcePolicy,
            draft.DesignBrief.Brief.Assumptions.ToArray(),
            assetAudit);
        return new VisualDeckBrandProfileBinding(
            new BrandProfileReference(
                draft.DesignBrief.Profile.Id,
                draft.DesignBrief.Profile.Version,
                draft.DesignBrief.Profile.ContentHash),
            draft.DesignBrief.StyleDirection.Id,
            slideBindings,
            audit);
    }

    private static bool MatchesActiveDraft(
        DraftState active,
        string title,
        int expectedSlideCount,
        VisualThemeSpec? theme,
        string? subject,
        string language,
        VisualDesignSpec? design,
        string templateSourceFileId,
        string templateLayoutId,
        ValidatedDesignBriefBinding? designBrief) =>
        string.Equals(active.Title, title, StringComparison.Ordinal)
        && active.ExpectedSlideCount == expectedSlideCount
        && ThemesEqual(active.Theme, theme)
        && string.Equals(active.Subject, subject, StringComparison.Ordinal)
        && string.Equals(active.Language, language, StringComparison.Ordinal)
        && Equals(active.Design, design)
        && string.Equals(
            active.TemplateSourceFileId,
            NormalizeTemplateSource(templateSourceFileId),
            StringComparison.OrdinalIgnoreCase)
        && string.Equals(
            active.TemplateLayoutId,
            NormalizeTemplateLayout(templateSourceFileId, templateLayoutId),
            StringComparison.Ordinal)
        && string.Equals(active.DesignBrief?.BriefId, designBrief?.BriefId, StringComparison.Ordinal)
        && string.Equals(
            active.DesignBrief?.Profile.ContentHash,
            designBrief?.Profile.ContentHash,
            StringComparison.Ordinal);

    private static bool ThemesEqual(VisualThemeSpec? first, VisualThemeSpec? second)
    {
        if (ReferenceEquals(first, second))
        {
            return true;
        }

        if (first is null || second is null)
        {
            return false;
        }

        return string.Equals(first.Preset, second.Preset, StringComparison.Ordinal)
            && string.Equals(first.PrimaryColor, second.PrimaryColor, StringComparison.Ordinal)
            && string.Equals(first.SecondaryColor, second.SecondaryColor, StringComparison.Ordinal)
            && string.Equals(first.AccentColor, second.AccentColor, StringComparison.Ordinal)
            && string.Equals(first.BackgroundColor, second.BackgroundColor, StringComparison.Ordinal)
            && string.Equals(first.TextColor, second.TextColor, StringComparison.Ordinal)
            && string.Equals(first.FontFace, second.FontFace, StringComparison.Ordinal)
            && string.Equals(first.HeadingFontFace, second.HeadingFontFace, StringComparison.Ordinal)
            && string.Equals(first.BodyFontFace, second.BodyFontFace, StringComparison.Ordinal)
            && string.Equals(first.SurfaceColor, second.SurfaceColor, StringComparison.Ordinal)
            && string.Equals(first.MutedTextColor, second.MutedTextColor, StringComparison.Ordinal)
            && string.Equals(first.PositiveColor, second.PositiveColor, StringComparison.Ordinal)
            && string.Equals(first.WarningColor, second.WarningColor, StringComparison.Ordinal)
            && string.Equals(first.CriticalColor, second.CriticalColor, StringComparison.Ordinal)
            && (first.DataSeriesColors ?? []).SequenceEqual(
                second.DataSeriesColors ?? [],
                StringComparer.Ordinal);
    }

    private static void ValidateBrandRecipes(
        DraftState draft,
        IReadOnlyList<VisualSlideSpec> slides,
        int startSlideNumber)
    {
        if (draft.DesignBrief is null)
        {
            return;
        }

        for (var index = 0; index < slides.Count; index++)
        {
            var slideNumber = startSlideNumber + index;
            var slide = slides[index];
            if (!draft.DesignBrief.AssetPlan.TryGetValue(slideNumber, out var plan))
            {
                throw new PptxValidationException(
                    "visual_slide_asset_plan_missing",
                    $"Slide {slideNumber} has no entry in the validated Asset Plan.");
            }

            if (!string.Equals(slide.RecipeId, plan.RecipeId, StringComparison.Ordinal))
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_mismatch",
                    $"Slide {slideNumber}.recipeId must be {plan.RecipeId}, as fixed by the validated Design Brief.");
            }

            var recipe = draft.DesignBrief.Profile.LayoutRecipes.Single(item =>
                string.Equals(item.Id, plan.RecipeId, StringComparison.Ordinal));
            if (slide.Kind != recipe.SemanticKind)
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_kind_mismatch",
                    $"Slide {slideNumber}.kind must be {recipe.SemanticKind} for recipe {recipe.Id}.");
            }

            var effectiveDensity = slide.Density
                ?? draft.Design?.Density
                ?? "balanced";
            if (!string.Equals(effectiveDensity, recipe.Density, StringComparison.OrdinalIgnoreCase))
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_density_mismatch",
                    $"Slide {slideNumber} effective density must be {recipe.Density} for recipe {recipe.Id}. Set the slide density explicitly when it differs from the brief base density.");
            }

            if (!string.Equals(slide.Variant, recipe.Variant, StringComparison.OrdinalIgnoreCase))
            {
                throw new PptxValidationException(
                    "visual_slide_recipe_variant_mismatch",
                    $"Slide {slideNumber}.variant must be {recipe.Variant} for recipe {recipe.Id}.");
            }
        }
    }

    private static VisualDeckDraftView CreateView(DraftState draft)
    {
        var accepted = draft.Slides.Count;
        var remaining = draft.ExpectedSlideCount - accepted;
        var instruction = remaining == 0
            ? "All requested slides are accepted. Call the appropriate pptx_finish_* tool once with this draft_id."
            : $"Call pptx_add_visual_slides_to_draft with this draft_id and the next 1-{Math.Min(MaximumBatchSlides, remaining)} slides. Omit startSlideNumber; the server will append at slide {accepted + 1}.";
        return new VisualDeckDraftView(
            draft.Id,
            "draft",
            draft.ExpectedSlideCount,
            accepted,
            accepted + 1,
            remaining,
            MaximumBatchSlides,
            draft.TemplateSourceFileId,
            draft.TemplateLayoutId,
            true,
            instruction);
    }

    private static void ValidateTemplateSelection(string sourceFileId, string layoutId)
    {
        if (string.IsNullOrWhiteSpace(sourceFileId) || sourceFileId.Length > 512)
        {
            throw new PptxValidationException(
                "visual_template_source_invalid",
                "templateSourceFileId must be default, none, latest, or an uploaded PPTX file_id.");
        }

        if (string.IsNullOrWhiteSpace(layoutId) || layoutId.Length > 512)
        {
            throw new PptxValidationException(
                "visual_template_layout_invalid",
                "templateLayoutId must be auto or an exact blank layout_id returned by pptx_analyze.");
        }
    }

    private static void EnsureTemplateSelectionMatches(
        DraftState draft,
        string? requestedSourceFileId,
        string? requestedLayoutId)
    {
        var requestedSource = requestedSourceFileId is null
            ? draft.TemplateSourceFileId
            : NormalizeTemplateSource(requestedSourceFileId);
        var requestedLayout = requestedLayoutId is null
            ? draft.TemplateLayoutId
            : NormalizeTemplateLayout(requestedSource, requestedLayoutId);
        if (!string.Equals(requestedSource, draft.TemplateSourceFileId, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(requestedLayout, draft.TemplateLayoutId, StringComparison.Ordinal))
        {
            throw new PptxValidationException(
                "visual_creative_direction_locked",
                "Template and theme choices are locked when pptx_start_visual_deck succeeds. Finish this draft with the locked choices; do not rebuild the whole deck.");
        }
    }

    private static string NormalizeTemplateSource(string sourceFileId) =>
        sourceFileId.Trim();

    private static string NormalizeTemplateLayout(string sourceFileId, string layoutId) =>
        string.Equals(sourceFileId, "none", StringComparison.OrdinalIgnoreCase)
            ? "auto"
            : layoutId.Trim();

    private void PruneExpired()
    {
        var now = timeProvider.GetUtcNow();
        foreach (var pair in drafts)
        {
            if (pair.Value.ExpiresAt <= now)
            {
                drafts.TryRemove(pair.Key, out _);
            }
        }
    }

    private static bool IsValidId(string value) =>
        value is not null
        && value.Length == 32
        && value.All(Uri.IsHexDigit);

    private enum DraftStatus
    {
        Collecting,
        Submitting,
        Submitted,
    }

    private sealed record DraftState(
        string Id,
        string UserScope,
        string ConversationScope,
        string Title,
        int ExpectedSlideCount,
        IReadOnlyList<VisualSlideSpec> Slides,
        VisualThemeSpec? Theme,
        string? Subject,
        string Language,
        VisualDesignSpec? Design,
        string TemplateSourceFileId,
        string TemplateLayoutId,
        ValidatedDesignBriefBinding? DesignBrief,
        DraftStatus Status,
        string? JobId,
        DateTimeOffset CreatedAt,
        DateTimeOffset ExpiresAt);
}

public sealed record VisualDeckDraftSubmission(
    VisualDeckSpec? Deck,
    string? ExistingJobId,
    string TemplateSourceFileId,
    string TemplateLayoutId);
