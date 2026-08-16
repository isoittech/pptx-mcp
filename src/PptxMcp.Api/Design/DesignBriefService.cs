using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Design;

public sealed partial class DesignBriefService(
    IOptions<PptxMcpOptions> options,
    TimeProvider timeProvider,
    BrandProfileCatalog catalog)
{
    private const int MaximumActiveBriefs = 256;
    private const int MaximumStyleAlternatives = 2;
    private const int MaximumPendingChoicesPerUser = 8;
    private static readonly HashSet<string> SupportedDensities = new(StringComparer.OrdinalIgnoreCase)
    {
        "airy",
        "balanced",
        "detailed",
    };

    private static readonly HashSet<string> SupportedCropIntents = new(StringComparer.OrdinalIgnoreCase)
    {
        "contain",
        "cover",
        "focalCenter",
        "focalLeft",
        "focalRight",
        "none",
    };

    private static readonly HashSet<string> SupportedAspectRatios = new(StringComparer.OrdinalIgnoreCase)
    {
        "landscape16x9",
        "landscape4x3",
        "square1x1",
        "portrait4x5",
        "flexible",
    };

    private static readonly HashSet<string> SupportedTextSafeAreas = new(StringComparer.OrdinalIgnoreCase)
    {
        "none",
        "left",
        "right",
        "top",
        "bottom",
    };

    private readonly PptxMcpOptions options = options.Value;
    private readonly ConcurrentDictionary<string, ValidatedDesignBriefBinding> briefs = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, PreparedDesignBriefSelection> preparedSelections = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, BriefOwnership> ownership = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, byte> startedBriefs = new(StringComparer.Ordinal);
    private readonly HashSet<string> startReservations = new(StringComparer.Ordinal);
    private readonly HashSet<BriefOwnership> conversationStartReservations = [];
    private readonly object stateGate = new();

    public DesignBriefValidationView Validate(
        CallerContext caller,
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan)
    {
        ArgumentNullException.ThrowIfNull(caller);
        ArgumentNullException.ThrowIfNull(brief);
        ArgumentNullException.ThrowIfNull(assetPlan);
        lock (stateGate)
        {
            PruneExpiredUnsafe();
            if (HasStartReservationForCallerUnsafe(caller))
            {
                throw new PptxValidationException(
                    "design_brief_start_in_progress",
                    "A Visual Deck start is in progress for this conversation. Wait for that call to finish before validating another Design Brief.");
            }

            if (FindBlockingSelectionUnsafe(caller) is not null)
            {
                throw new PptxValidationException(
                    "design_brief_choice_pending",
                    "A Design Brief card is awaiting a user choice in this conversation. Apply that server-issued choice before validating or starting another brief.");
            }
        }

        var expiresAt = timeProvider.GetUtcNow().AddMinutes(options.DesignBriefLifetimeMinutes);
        while (true)
        {
            var briefId = NewOpaqueId();
            var binding = ValidateCandidate(
                briefId,
                brief,
                assetPlan,
                expiresAt,
                DesignBriefSelectionSource.AgentDefault);
            lock (stateGate)
            {
                PruneExpiredUnsafe();
                if (HasStartReservationForCallerUnsafe(caller))
                {
                    throw new PptxValidationException(
                        "design_brief_start_in_progress",
                        "A Visual Deck start is in progress for this conversation. Wait for that call to finish before validating another Design Brief.");
                }

                if (FindBlockingSelectionUnsafe(caller) is not null)
                {
                    throw new PptxValidationException(
                        "design_brief_choice_pending",
                        "A Design Brief card is awaiting a user choice in this conversation. Apply that server-issued choice before validating or starting another brief.");
                }

                EnsureCapacityUnsafe();
                if (ownership.TryAdd(briefId, new BriefOwnership(caller.UserScope, caller.ConversationScope))
                    && briefs.TryAdd(briefId, binding))
                {
                    return CreateValidationView(binding);
                }

                briefs.TryRemove(briefId, out _);
                ownership.TryRemove(briefId, out _);
            }
        }
    }

    internal PreparedDesignBriefCard Prepare(
        CallerContext caller,
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan,
        IReadOnlyList<DesignBriefStyleAlternative>? alternatives,
        IReadOnlyList<DesignBriefNoPhotoOverride>? noPhotoOverrides,
        bool replacePendingChoice = false)
    {
        ArgumentNullException.ThrowIfNull(caller);
        ArgumentNullException.ThrowIfNull(brief);
        ArgumentNullException.ThrowIfNull(assetPlan);
        var requestedAlternatives = alternatives ?? [];
        var requestedNoPhotoOverrides = noPhotoOverrides ?? [];
        if (requestedAlternatives.Count > MaximumStyleAlternatives
            || requestedAlternatives.Any(static alternative => alternative is null))
        {
            throw new PptxValidationException(
                "design_brief_alternatives_invalid",
                $"alternatives must contain no more than {MaximumStyleAlternatives} compact style choices.");
        }

        if (requestedNoPhotoOverrides.Count > 0 && requestedAlternatives.Count > 1)
        {
            throw new PptxValidationException(
                "design_brief_card_option_limit_exceeded",
                "A Design Brief card supports at most three choices including the recommendation. When noPhotoOverrides are supplied, include at most one style alternative.");
        }

        if (requestedNoPhotoOverrides.Any(static item => item is null))
        {
            throw new PptxValidationException(
                "design_brief_no_photo_invalid",
                "noPhotoOverrides must not contain null items.");
        }

        var duplicateDirections = requestedAlternatives
            .Select(static alternative => alternative.StyleDirectionId)
            .Append(brief.StyleDirectionId)
            .GroupBy(static value => value, StringComparer.Ordinal)
            .Any(static group => group.Count() > 1);
        if (duplicateDirections)
        {
            throw new PptxValidationException(
                "design_brief_alternatives_invalid",
                "Every alternative must select a distinct style_direction_id different from the recommended direction.");
        }

        var expiresAt = timeProvider.GetUtcNow().AddMinutes(options.DesignBriefLifetimeMinutes);
        while (true)
        {
            var briefId = NewOpaqueId();
            var recommended = CreateCardOption(
                DesignBriefCardOptionKind.Recommended,
                ValidateCandidate(
                    briefId,
                    brief,
                    assetPlan,
                    expiresAt,
                    DesignBriefSelectionSource.UserCard));
            var alternativeOptions = requestedAlternatives
                .Select(alternative => CreateAlternativeOption(
                    briefId,
                    brief,
                    assetPlan,
                    alternative,
                    expiresAt))
                .ToArray();
            var noPhoto = CreateNoPhotoOption(
                briefId,
                brief,
                assetPlan,
                requestedNoPhotoOverrides,
                expiresAt);
            var allOptions = new List<PreparedDesignBriefOption>(1 + alternativeOptions.Length + 1)
            {
                recommended,
            };
            allOptions.AddRange(alternativeOptions);
            if (noPhoto is not null)
            {
                allOptions.Add(noPhoto);
            }

            if (allOptions.Count < 2)
            {
                throw new PptxValidationException(
                    "design_brief_card_choice_required",
                    "A Design Brief card must offer at least two executable choices with a material visual difference. Use pptx_validate_design_brief when only one safe direction exists.");
            }

            EnsureMateriallyDistinctOptions(allOptions);

            var optionsById = allOptions.ToDictionary(
                static option => option.OptionId,
                StringComparer.Ordinal);
            var choiceSessionId = NewOpaqueId();
            var selection = new PreparedDesignBriefSelection(
                choiceSessionId,
                briefId,
                expiresAt,
                recommended.Binding.Profile.Id,
                recommended.Binding.Profile.Version,
                recommended.Binding.Profile.ContentHash,
                optionsById,
                null,
                false,
                false);

            lock (stateGate)
            {
                PruneExpiredUnsafe();
                if (HasStartReservationForCallerUnsafe(caller))
                {
                    throw new PptxValidationException(
                        "design_brief_start_in_progress",
                        "A Visual Deck start is in progress for this conversation. Wait for that call to finish before preparing another Design Brief card.");
                }

                var existingPending = FindBlockingSelectionUnsafe(caller);
                if (existingPending is not null && !replacePendingChoice)
                {
                    throw new PptxValidationException(
                        "design_brief_choice_pending",
                        "A Design Brief card is already awaiting a user choice in this conversation. Reuse it, or explicitly replace it only after the user asks to replace the card.");
                }

                if (existingPending is not null)
                {
                    if (existingPending.SelectedOptionId is not null)
                    {
                        throw new PptxValidationException(
                            "design_brief_choice_already_applied",
                            "The existing Design Brief card already has an applied choice. It cannot be replaced before that selected start succeeds or the choice expires.");
                    }

                    preparedSelections.TryRemove(existingPending.ChoiceSessionId, out _);
                    ownership.TryRemove(existingPending.ChoiceSessionId, out _);
                }

                EnsureCapacityUnsafe();
                EnsurePerUserCapacityUnsafe(caller);
                if (ownership.TryAdd(choiceSessionId, new BriefOwnership(caller.UserScope, caller.ConversationScope))
                    && preparedSelections.TryAdd(choiceSessionId, selection))
                {
                    InvalidateUnstartedBriefsUnsafe(caller);
                    return new PreparedDesignBriefCard(
                        choiceSessionId,
                        expiresAt,
                        recommended,
                        alternativeOptions,
                        noPhoto);
                }

                preparedSelections.TryRemove(choiceSessionId, out _);
                ownership.TryRemove(choiceSessionId, out _);
            }
        }
    }

    public DesignBriefValidationView ApplyCardSelection(
        CallerContext caller,
        string choiceSessionId,
        string optionId)
    {
        ArgumentNullException.ThrowIfNull(caller);
        ValidateOpaqueId(choiceSessionId, "choiceSessionId");
        ValidateOpaqueId(optionId, "optionId");
        choiceSessionId = choiceSessionId.ToLowerInvariant();
        optionId = optionId.ToLowerInvariant();

        lock (stateGate)
        {
            if (!ownership.TryGetValue(choiceSessionId, out var owner)
                || !string.Equals(owner.UserScope, caller.UserScope, StringComparison.Ordinal)
                || !string.Equals(owner.ConversationScope, caller.ConversationScope, StringComparison.Ordinal)
                || !preparedSelections.TryGetValue(choiceSessionId, out var selection))
            {
                PruneExpiredUnsafe();
                throw new PptxValidationException(
                    "design_brief_action_not_found",
                    "The Design Brief choice session was not found for this user and conversation. Prepare a new card in this conversation.");
            }

            if (selection.StartReserved)
            {
                throw new PptxValidationException(
                    "design_brief_action_start_in_progress",
                    "This Design Brief choice is already starting a Visual Deck. Ignore duplicate card actions until the start call completes.");
            }

            if (selection.ExpiresAt <= timeProvider.GetUtcNow())
            {
                preparedSelections.TryRemove(choiceSessionId, out _);
                ownership.TryRemove(choiceSessionId, out _);
                throw new PptxValidationException(
                    "design_brief_action_expired",
                    "The Design Brief choice expired. Refresh the catalog and prepare a new card.");
            }

            PruneExpiredUnsafe();

            if (selection.SelectedOptionId is { } selectedOptionId)
            {
                if (selection.StartCompleted)
                {
                    throw new PptxValidationException(
                        "design_brief_action_already_started",
                        "This Design Brief choice already started a Visual Deck. Ignore the stale card action and continue the existing draft.");
                }

                if (!string.Equals(selectedOptionId, optionId, StringComparison.Ordinal))
                {
                    throw new PptxValidationException(
                        "design_brief_action_replayed",
                        "This Design Brief card already selected a different option. Do not change it by replaying the embedded action.");
                }

                if (!briefs.TryGetValue(selection.BriefId, out var existing))
                {
                    throw new PptxValidationException(
                        "design_brief_action_state_invalid",
                        "The selected Design Brief state is incomplete. Prepare a new card.");
                }

                return CreateValidationView(existing);
            }

            if (!selection.Options.TryGetValue(optionId, out var selected))
            {
                throw new PptxValidationException(
                    "design_brief_action_tampered",
                    "The Design Brief option is not one of the server-issued choices for this card.");
            }

            var currentProfile = catalog.GetSnapshot(selected.Binding.Brief.BrandProfile);
            if (!string.Equals(currentProfile.Id, selection.ProfileId, StringComparison.Ordinal)
                || !string.Equals(currentProfile.Version, selection.ProfileVersion, StringComparison.Ordinal)
                || !string.Equals(currentProfile.ContentHash, selection.ProfileContentHash, StringComparison.Ordinal))
            {
                throw new PptxValidationException(
                    "brand_profile_version_mismatch",
                    "The Brand Profile binding for this Design Brief card changed. Prepare a new card.");
            }

            if (briefs.ContainsKey(selection.BriefId)
                || ownership.ContainsKey(selection.BriefId)
                || !ownership.TryAdd(
                    selection.BriefId,
                    new BriefOwnership(caller.UserScope, caller.ConversationScope)))
            {
                throw new PptxValidationException(
                    "design_brief_action_replayed",
                    "This Design Brief card has already been applied.");
            }

            if (!briefs.TryAdd(selection.BriefId, selected.Binding))
            {
                ownership.TryRemove(selection.BriefId, out _);
                throw new PptxValidationException(
                    "design_brief_action_replayed",
                    "This Design Brief card has already been applied.");
            }

            preparedSelections[choiceSessionId] = selection with { SelectedOptionId = optionId };
            return CreateValidationView(selected.Binding);
        }
    }

    internal void DiscardPrepared(
        CallerContext caller,
        string choiceSessionId)
    {
        ArgumentNullException.ThrowIfNull(caller);
        ValidateOpaqueId(choiceSessionId, "choiceSessionId");
        choiceSessionId = choiceSessionId.ToLowerInvariant();
        lock (stateGate)
        {
            if (!IsOwnedByCallerUnsafe(choiceSessionId, caller)
                || !preparedSelections.TryGetValue(choiceSessionId, out var selection)
                || selection.SelectedOptionId is not null
                || selection.StartReserved)
            {
                return;
            }

            preparedSelections.TryRemove(choiceSessionId, out _);
            ownership.TryRemove(choiceSessionId, out _);
        }
    }

    public DesignBriefSelectionCancellationView CancelPendingSelection(CallerContext caller)
    {
        ArgumentNullException.ThrowIfNull(caller);
        lock (stateGate)
        {
            PruneExpiredUnsafe();
            var active = FindBlockingSelectionUnsafe(caller);
            var selection = active ?? preparedSelections.Values
                .Where(item => IsOwnedByCallerUnsafe(item.ChoiceSessionId, caller))
                .OrderByDescending(static item => item.ExpiresAt)
                .FirstOrDefault();
            if (selection is null)
            {
                throw new PptxValidationException(
                    "design_brief_choice_not_pending",
                    "This conversation has no pending Design Brief card to cancel.");
            }

            if (active is null
                || selection.SelectedOptionId is not null
                || selection.StartReserved
                || selection.StartCompleted)
            {
                throw new PptxValidationException(
                    "design_brief_choice_cancel_forbidden",
                    "The Design Brief choice is already selected or starting and cannot be cancelled. Continue only the selected brief or existing draft.");
            }

            preparedSelections.TryRemove(selection.ChoiceSessionId, out _);
            ownership.TryRemove(selection.ChoiceSessionId, out _);
            return new DesignBriefSelectionCancellationView(
                "cancelled",
                "The unselected Design Brief card was cancelled. The caller may now validate the safe recommendation directly or prepare a new materially different card.");
        }
    }

    public ValidatedDesignBriefBinding? AuthorizeForStart(
        CallerContext caller,
        string? briefId,
        int expectedSlideCount,
        string templateSourceFileId,
        VisualThemeSpec? requestedTheme,
        VisualDesignSpec? requestedDesign,
        bool userRequestedNewWorkflow = false)
    {
        ArgumentNullException.ThrowIfNull(caller);
        if (string.IsNullOrWhiteSpace(briefId))
        {
            lock (stateGate)
            {
                PruneExpiredUnsafe();
                if (FindBlockingSelectionUnsafe(caller) is { } selection)
                {
                    throw new PptxValidationException(
                        selection.SelectedOptionId is null
                            ? "design_brief_not_confirmed"
                            : "design_brief_selection_required",
                        selection.SelectedOptionId is null
                            ? "A Design Brief card in this conversation is still pending. Apply its server-issued option before starting any Visual Deck."
                            : "This conversation has a selected Design Brief card. Start only the brief_id returned by pptx_apply_design_brief_action.");
                }
            }

            if (options.RequireDesignBrief)
            {
                throw new PptxValidationException(
                    "design_brief_required",
                    "This deployment requires a validated Design Brief. Call pptx_get_design_catalog and pptx_validate_design_brief before pptx_start_visual_deck.");
            }

            return null;
        }

        if (!OpaqueIdRegex().IsMatch(briefId))
        {
            throw new PptxValidationException(
                "design_brief_not_found",
                "The Design Brief was not found for this user and conversation. Validate a new brief in this conversation.");
        }

        briefId = briefId.ToLowerInvariant();
        lock (stateGate)
        {
            var now = timeProvider.GetUtcNow();
            var callerOwnsBrief = briefs.TryGetValue(briefId, out var brief)
                && ownership.TryGetValue(briefId, out var owner)
                && string.Equals(owner.UserScope, caller.UserScope, StringComparison.Ordinal)
                && string.Equals(owner.ConversationScope, caller.ConversationScope, StringComparison.Ordinal);

            // Classify this exact caller-owned ID before pruning unrelated state. Otherwise an
            // expired direct brief disappears first and is reported as an ambiguous not-found ID.
            if (callerOwnsBrief && brief!.ExpiresAt <= now)
            {
                briefs.TryRemove(brief.BriefId, out _);
                ownership.TryRemove(brief.BriefId, out _);
                startedBriefs.TryRemove(brief.BriefId, out _);
                PruneExpiredUnsafe();
                throw new PptxValidationException(
                    "design_brief_expired",
                    "The Design Brief expired. Refresh the design catalog and validate the brief again.");
            }

            PruneExpiredUnsafe();
            var selectionGate = FindBlockingSelectionUnsafe(caller);
            if (selectionGate?.SelectedOptionId is null && selectionGate is not null)
            {
                throw new PptxValidationException(
                    "design_brief_not_confirmed",
                    "A Design Brief card in this conversation is still pending. Apply its server-issued option before starting any Design Brief.");
            }

            if (selectionGate is not null
                && !string.Equals(selectionGate.BriefId, briefId, StringComparison.Ordinal))
            {
                throw new PptxValidationException(
                    "design_brief_selection_superseded",
                    "The user selected a different Design Brief from the card. Start only the brief_id returned by pptx_apply_design_brief_action.");
            }

            if (!callerOwnsBrief)
            {
                throw new PptxValidationException(
                    "design_brief_not_found",
                    "The Design Brief was not found for this user and conversation. Validate a new brief in this conversation.");
            }

            var authorizedBrief = brief!;
            if (userRequestedNewWorkflow && startedBriefs.ContainsKey(authorizedBrief.BriefId))
            {
                throw new PptxValidationException(
                    "design_brief_already_started",
                    "A new Visual Deck workflow requires a newly validated or newly selected Design Brief. Do not reuse a brief that already started a deck.");
            }

            if (authorizedBrief.Brief.ExpectedSlideCount != expectedSlideCount)
            {
                throw new PptxValidationException(
                    "design_brief_slide_count_mismatch",
                    $"expectedSlideCount must remain {authorizedBrief.Brief.ExpectedSlideCount}, as finalized in the Design Brief.");
            }

            if (!string.Equals(authorizedBrief.Profile.TemplateSource, templateSourceFileId.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                throw new PptxValidationException(
                    "design_brief_template_mismatch",
                    $"templateSourceFileId must remain {authorizedBrief.Profile.TemplateSource}, as fixed by the Brand Profile.");
            }

            if (requestedTheme is not null || requestedDesign is not null)
            {
                throw new PptxValidationException(
                    "design_brief_creative_direction_conflict",
                    "Omit theme and design when using briefId; the validated Brand Profile version and style direction supply immutable values.");
            }

            return authorizedBrief;
        }
    }

    public bool ReserveStart(
        CallerContext caller,
        ValidatedDesignBriefBinding? brief)
    {
        ArgumentNullException.ThrowIfNull(caller);
        lock (stateGate)
        {
            PruneExpiredUnsafe();
            var ownerKey = new BriefOwnership(caller.UserScope, caller.ConversationScope);
            if (conversationStartReservations.Contains(ownerKey))
            {
                throw new PptxValidationException(
                    "design_brief_start_in_progress",
                    "A Visual Deck start is already in progress for this conversation.");
            }

            var selection = FindBlockingSelectionUnsafe(caller);
            if (brief is null)
            {
                if (selection is not null)
                {
                    throw new PptxValidationException(
                        selection.SelectedOptionId is null
                            ? "design_brief_not_confirmed"
                            : "design_brief_selection_required",
                        "A Design Brief card now controls this conversation. Apply and start its selected brief instead of an unbound deck.");
                }

                conversationStartReservations.Add(ownerKey);
                return true;
            }

            if (!briefs.TryGetValue(brief.BriefId, out var current)
                || !ReferenceEquals(current, brief)
                || !IsOwnedByCallerUnsafe(brief.BriefId, caller))
            {
                throw new PptxValidationException(
                    "design_brief_start_state_invalid",
                    "The Design Brief changed before start could be reserved. Validate or apply the choice again.");
            }

            if (selection is not null
                && (selection.SelectedOptionId is null
                    || !string.Equals(selection.BriefId, brief.BriefId, StringComparison.Ordinal)))
            {
                throw new PptxValidationException(
                    "design_brief_start_state_changed",
                    "The Design Brief card state changed before generation began. Start only the currently selected brief.");
            }

            if (startedBriefs.ContainsKey(brief.BriefId))
            {
                conversationStartReservations.Add(ownerKey);
                return true;
            }

            if (!startReservations.Add(brief.BriefId))
            {
                throw new PptxValidationException(
                    "design_brief_start_in_progress",
                    "This Design Brief is already being used to start a Visual Deck in this conversation.");
            }

            conversationStartReservations.Add(ownerKey);

            if (selection is not null)
            {
                preparedSelections[selection.ChoiceSessionId] = selection with { StartReserved = true };
            }

            return true;
        }
    }

    public void ReleaseStartReservation(
        CallerContext caller,
        ValidatedDesignBriefBinding? brief)
    {
        ArgumentNullException.ThrowIfNull(caller);
        lock (stateGate)
        {
            conversationStartReservations.Remove(
                new BriefOwnership(caller.UserScope, caller.ConversationScope));
            if (brief is null)
            {
                return;
            }

            if (!IsOwnedByCallerUnsafe(brief.BriefId, caller))
            {
                return;
            }

            startReservations.Remove(brief.BriefId);
            var selection = preparedSelections.Values.FirstOrDefault(item =>
                item.StartReserved
                && !item.StartCompleted
                && string.Equals(item.BriefId, brief.BriefId, StringComparison.Ordinal)
                && IsOwnedByCallerUnsafe(item.ChoiceSessionId, caller));
            if (selection is not null)
            {
                preparedSelections[selection.ChoiceSessionId] = selection with { StartReserved = false };
            }
        }
    }

    public void MarkStartSucceeded(CallerContext caller, string? briefId)
    {
        ArgumentNullException.ThrowIfNull(caller);
        lock (stateGate)
        {
            if (string.IsNullOrWhiteSpace(briefId) || !OpaqueIdRegex().IsMatch(briefId))
            {
                conversationStartReservations.Remove(
                    new BriefOwnership(caller.UserScope, caller.ConversationScope));
                return;
            }

            briefId = briefId.ToLowerInvariant();
            if (startedBriefs.ContainsKey(briefId))
            {
                startReservations.Remove(briefId);
                conversationStartReservations.Remove(
                    new BriefOwnership(caller.UserScope, caller.ConversationScope));
                return;
            }

            if (!briefs.ContainsKey(briefId)
                || !ownership.TryGetValue(briefId, out var briefOwner)
                || !string.Equals(briefOwner.UserScope, caller.UserScope, StringComparison.Ordinal)
                || !string.Equals(briefOwner.ConversationScope, caller.ConversationScope, StringComparison.Ordinal)
                || !startReservations.Remove(briefId))
            {
                throw new PptxValidationException(
                    "design_brief_start_state_invalid",
                    "The Design Brief start state no longer belongs to this user and conversation.");
            }

            startedBriefs[briefId] = 0;
            var selection = preparedSelections.Values.FirstOrDefault(item =>
                !item.StartCompleted
                && string.Equals(item.BriefId, briefId, StringComparison.Ordinal)
                && IsOwnedByCallerUnsafe(item.ChoiceSessionId, caller));
            if (selection is not null)
            {
                preparedSelections[selection.ChoiceSessionId] = selection with
                {
                    StartReserved = false,
                    StartCompleted = true,
                };
            }

            conversationStartReservations.Remove(
                new BriefOwnership(caller.UserScope, caller.ConversationScope));
        }
    }

    private ValidatedDesignBriefBinding ValidateCandidate(
        string briefId,
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan,
        DateTimeOffset expiresAt,
        DesignBriefSelectionSource selectionSource)
    {
        ValidateText(brief.Audience, "brief.audience", 1, 240);
        ValidateText(brief.Purpose, "brief.purpose", 1, 500);
        ValidateText(brief.DesiredTone, "brief.desired_tone", 1, 160);
        ValidateText(brief.VisualStrategy, "brief.visual_strategy", 1, 800);
        ValidateIdentifier(brief.StyleDirectionId, "brief.style_direction_id");
        if (!SupportedDensities.Contains(brief.Density))
        {
            throw new PptxValidationException(
                "design_brief_density_invalid",
                "brief.density must be airy, balanced, or detailed.");
        }

        if (brief.ExpectedSlideCount is < 1 || brief.ExpectedSlideCount > options.MaxSlides)
        {
            throw new PptxValidationException(
                "design_brief_slide_count_invalid",
                $"brief.expected_slide_count must be between 1 and {options.MaxSlides}.");
        }

        if (brief.Assumptions is null || brief.Assumptions.Count > 16)
        {
            throw new PptxValidationException(
                "design_brief_assumptions_invalid",
                "brief.assumptions must contain no more than 16 concise items.");
        }

        for (var index = 0; index < brief.Assumptions.Count; index++)
        {
            var assumption = brief.Assumptions[index];
            if (assumption is null)
            {
                throw new PptxValidationException(
                    "design_brief_assumptions_invalid",
                    $"brief.assumptions[{index}] must not be null.");
            }

            ValidateText(assumption.Text, $"brief.assumptions[{index}].text", 1, 300);
            if (assumption.Status == DesignAssumptionStatus.NeedsConfirmation)
            {
                throw new PptxValidationException(
                    "design_brief_confirmation_required",
                    $"brief.assumptions[{index}] still needs confirmation. Ask the user or select a safe explicit fallback before validation.");
            }
        }

        if (brief.QuestionsForUser is null || brief.QuestionsForUser.Count > 0)
        {
            throw new PptxValidationException(
                "design_brief_questions_unresolved",
                "brief.questions_for_user must be empty. Resolve material questions or record a safe explicit fallback before validating the brief.");
        }

        if (brief.BrandProfile is null)
        {
            throw new PptxValidationException(
                "brand_profile_reference_required",
                "brief.brand_profile must copy id, version, and content_hash from pptx_get_design_catalog.");
        }

        var profile = catalog.GetSnapshot(brief.BrandProfile);
        var styleDirection = profile.Detail.StyleDirections.SingleOrDefault(direction =>
            string.Equals(direction.Id, brief.StyleDirectionId, StringComparison.Ordinal));
        if (styleDirection is null)
        {
            throw new PptxValidationException(
                "brand_style_direction_not_found",
                "brief.style_direction_id was not found in the selected immutable Brand Profile.");
        }

        if (!styleDirection.SupportedDensities.Contains(brief.Density, StringComparer.OrdinalIgnoreCase))
        {
            throw new PptxValidationException(
                "design_brief_density_not_supported",
                "brief.density is not supported by the selected Brand Profile style direction.");
        }

        var planBySlide = ValidateAssetPlan(brief, assetPlan, profile, styleDirection);
        var theme = new VisualThemeSpec(
            styleDirection.ThemePreset,
            profile.Detail.ColorRoles.Primary,
            profile.Detail.ColorRoles.Secondary,
            profile.Detail.ColorRoles.Accent,
            profile.Detail.ColorRoles.Background,
            profile.Detail.ColorRoles.Text,
            FontFace: null,
            HeadingFontFace: profile.Detail.Typography.HeadingFont,
            BodyFontFace: profile.Detail.Typography.BodyFont,
            SurfaceColor: profile.Detail.ColorRoles.Surface,
            MutedTextColor: profile.Detail.ColorRoles.MutedText,
            PositiveColor: profile.Detail.ColorRoles.Positive,
            WarningColor: profile.Detail.ColorRoles.Warning,
            CriticalColor: profile.Detail.ColorRoles.Critical,
            DataSeriesColors: profile.Detail.ColorRoles.DataSeries);
        var design = new VisualDesignSpec(
            styleDirection.DesignStyle,
            brief.Density,
            styleDirection.Motif);
        return new ValidatedDesignBriefBinding(
            briefId,
            brief,
            profile,
            styleDirection,
            theme,
            design,
            planBySlide,
            expiresAt,
            selectionSource);
    }

    private PreparedDesignBriefOption CreateAlternativeOption(
        string briefId,
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan,
        DesignBriefStyleAlternative alternative,
        DateTimeOffset expiresAt)
    {
        if (alternative.RecipeIds is null)
        {
            throw new PptxValidationException(
                "design_brief_alternatives_invalid",
                "Each alternative must provide exactly one recipe_id per slide.");
        }

        ValidateIdentifier(alternative.StyleDirectionId, "alternatives.style_direction_id");
        if (!SupportedDensities.Contains(alternative.Density)
            || alternative.RecipeIds.Count != brief.ExpectedSlideCount)
        {
            throw new PptxValidationException(
                "design_brief_alternatives_invalid",
                "Each alternative must use a supported density and exactly one recipe_id per slide.");
        }

        var alternativeBrief = brief with
        {
            StyleDirectionId = alternative.StyleDirectionId,
            Density = alternative.Density,
        };
        var alternativePlan = assetPlan
            .Select((item, index) => item with { RecipeId = alternative.RecipeIds[index] })
            .ToArray();
        return CreateCardOption(
            DesignBriefCardOptionKind.Alternative,
            ValidateCandidate(
                briefId,
                alternativeBrief,
                alternativePlan,
                expiresAt,
                DesignBriefSelectionSource.UserCard));
    }

    private PreparedDesignBriefOption? CreateNoPhotoOption(
        string briefId,
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan,
        IReadOnlyList<DesignBriefNoPhotoOverride>? noPhotoOverrides,
        DateTimeOffset expiresAt)
    {
        var requested = noPhotoOverrides ?? [];
        if (requested.Count == 0)
        {
            return null;
        }

        var photoSlides = assetPlan
            .Where(static item => item.PreferredMedium == AssetPreferredMedium.Photo)
            .Select(static item => item.SlideNumber)
            .Order()
            .ToArray();
        var requestedSlides = requested
            .Select(static item => item.SlideNumber)
            .Order()
            .ToArray();
        if (photoSlides.Length == 0
            || !photoSlides.SequenceEqual(requestedSlides)
            || requested.Select(static item => item.SlideNumber).Distinct().Count() != requested.Count)
        {
            throw new PptxValidationException(
                "design_brief_no_photo_invalid",
                "noPhotoOverrides must replace every and only slide whose recommended preferred_medium is photo.");
        }

        var overrideBySlide = requested.ToDictionary(static item => item.SlideNumber);
        var noPhotoPlan = assetPlan
            .Select(item => overrideBySlide.TryGetValue(item.SlideNumber, out var replacement)
                ? CreateNoPhotoAssetPlanItem(item, replacement)
                : item)
            .ToArray();
        return CreateCardOption(
            DesignBriefCardOptionKind.NoPhoto,
            ValidateCandidate(
                briefId,
                brief,
                noPhotoPlan,
                expiresAt,
                DesignBriefSelectionSource.UserCard));
    }

    private static AssetPlanItem CreateNoPhotoAssetPlanItem(
        AssetPlanItem original,
        DesignBriefNoPhotoOverride replacement)
    {
        if (replacement.PreferredMedium == AssetPreferredMedium.Photo
            || replacement.Acquisition is not (AssetAcquisition.NativeDraw or AssetAcquisition.None)
            || replacement.Acquisition == AssetAcquisition.None
                && replacement.PreferredMedium != AssetPreferredMedium.None
            || replacement.Acquisition == AssetAcquisition.NativeDraw
                && replacement.PreferredMedium is not (
                    AssetPreferredMedium.Icon
                    or AssetPreferredMedium.NativeDiagram
                    or AssetPreferredMedium.Chart))
        {
            throw new PptxValidationException(
                "design_brief_no_photo_invalid",
                "A photo-free override must use nativeDraw with an editable medium, or the canonical no-asset combination.");
        }

        ValidateIdentifier(replacement.RecipeId, "noPhotoOverrides.recipe_id");
        return original with
        {
            RecipeId = replacement.RecipeId,
            PreferredMedium = replacement.PreferredMedium,
            Acquisition = replacement.Acquisition,
            Fallback = AssetFallback.None,
            Status = replacement.Acquisition == AssetAcquisition.NativeDraw
                ? AssetPlanStatus.Ready
                : AssetPlanStatus.Omitted,
            LicenseStatus = AssetLicenseStatus.NotRequired,
            AssetId = null,
            ApprovedAssetCollectionId = null,
            AttributionRef = null,
            CropIntent = null,
            AspectRatio = null,
            TextSafeArea = null,
            VisualObjectAssetIds = null,
        };
    }

    private static PreparedDesignBriefOption CreateCardOption(
        DesignBriefCardOptionKind kind,
        ValidatedDesignBriefBinding binding) =>
        new(NewOpaqueId(), kind, binding, CreateCardAssetSummary(binding.AssetPlan.Values));

    private static void EnsureMateriallyDistinctOptions(
        IReadOnlyList<PreparedDesignBriefOption> options)
    {
        var fingerprints = new Dictionary<string, PreparedDesignBriefOption>(StringComparer.Ordinal);
        foreach (var option in options)
        {
            var fingerprint = CreateEffectiveRenderFingerprint(option.Binding);
            if (fingerprints.TryGetValue(fingerprint, out var duplicate))
            {
                var errorCode = option.Kind == DesignBriefCardOptionKind.NoPhoto
                    ? "design_brief_no_photo_has_no_visual_difference"
                    : "design_brief_alternative_has_no_visual_difference";
                throw new PptxValidationException(
                    errorCode,
                    $"The {option.Kind} choice has the same effective theme, design, and per-slide render recipe as the {duplicate.Kind} choice. A different ID or Asset Plan label alone is not a visual alternative.");
            }

            fingerprints.Add(fingerprint, option);
        }
    }

    private static string CreateEffectiveRenderFingerprint(ValidatedDesignBriefBinding binding)
    {
        var theme = binding.Theme;
        var design = binding.Design;
        static string Token(string? value) => (value ?? string.Empty).Trim().ToLowerInvariant();
        static string Color(string? value) => Token(value).TrimStart('#');
        var parts = new List<string>
        {
            // A validated Brand Profile supplies every renderer color role explicitly. The
            // preset therefore contributes no effective value and cannot define a real choice.
            Color(theme.PrimaryColor),
            Color(theme.SecondaryColor),
            Color(theme.AccentColor),
            Color(theme.BackgroundColor),
            Color(theme.TextColor),
            Token(theme.FontFace),
            Token(theme.HeadingFontFace),
            Token(theme.BodyFontFace),
            Color(theme.SurfaceColor),
            Color(theme.MutedTextColor),
            Color(theme.PositiveColor),
            Color(theme.WarningColor),
            Color(theme.CriticalColor),
            string.Join(',', (theme.DataSeriesColors ?? []).Select(Color)),
            Token(design.Style),
            Token(design.Motif),
        };
        foreach (var pair in binding.AssetPlan.OrderBy(static pair => pair.Key))
        {
            var recipe = binding.Profile.LayoutRecipes.Single(item =>
                string.Equals(item.Id, pair.Value.RecipeId, StringComparison.Ordinal));
            parts.Add(pair.Key.ToString(System.Globalization.CultureInfo.InvariantCulture));
            parts.Add(Token(recipe.SemanticKind.ToString()));
            parts.Add(Token(recipe.Variant));
            parts.Add(Token(recipe.Density));
            parts.Add(recipe.ItemCount?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return string.Join('\u001f', parts);
    }

    private static DesignBriefCardAssetSummary CreateCardAssetSummary(
        IEnumerable<AssetPlanItem> assetPlan)
    {
        var items = assetPlan.ToArray();
        return new DesignBriefCardAssetSummary(
            items.Count(static item => item.Status == AssetPlanStatus.Ready),
            items.Count(static item => item.Status == AssetPlanStatus.NeedsUser),
            items.Count(static item => item.Status == AssetPlanStatus.FallbackSelected),
            items.Count(static item => item.Status == AssetPlanStatus.Omitted));
    }

    private static DesignBriefValidationView CreateValidationView(ValidatedDesignBriefBinding binding)
    {
        var values = binding.AssetPlan.Values;
        var summary = new AssetPlanSummary(
            values.Count(static item => item.Acquisition == AssetAcquisition.NativeDraw),
            values.Count(static item => item.Acquisition == AssetAcquisition.None),
            values.Count(static item => item.Status == AssetPlanStatus.FallbackSelected),
            values.Count(static item =>
                item.Acquisition == AssetAcquisition.UserUpload
                && item.Status == AssetPlanStatus.Ready));
        return new DesignBriefValidationView(
            binding.BriefId,
            "validated",
            binding.ExpiresAt,
            binding.Brief.BrandProfile,
            binding.Brief.StyleDirectionId,
            binding.Brief.ExpectedSlideCount,
            summary,
            "Call pptx_start_visual_deck with this brief_id, the same expectedSlideCount, and the profile template_source. Omit theme and design because the validated profile direction supplies them.");
    }

    private static Dictionary<int, AssetPlanItem> ValidateAssetPlan(
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan,
        BrandProfileSnapshot profile,
        BrandStyleDirection styleDirection)
    {
        if (assetPlan.Count != brief.ExpectedSlideCount
            || assetPlan.Any(static item => item is null))
        {
            throw new PptxValidationException(
                "asset_plan_slide_coverage_invalid",
                "assetPlan must contain exactly one complete item for every expected slide.");
        }

        var result = new Dictionary<int, AssetPlanItem>();
        for (var index = 0; index < assetPlan.Count; index++)
        {
            var item = assetPlan[index];
            var expectedSlideNumber = index + 1;
            if (item.SlideNumber != expectedSlideNumber || !result.TryAdd(item.SlideNumber, item))
            {
                throw new PptxValidationException(
                    "asset_plan_slide_order_invalid",
                    $"assetPlan[{index}].slide_number must be {expectedSlideNumber}; plans must be complete and ordered without duplicates.");
            }

            ValidateIdentifier(item.Purpose, $"assetPlan[{index}].purpose");
            ValidateIdentifier(item.RecipeId, $"assetPlan[{index}].recipe_id");
            var recipe = profile.LayoutRecipes.SingleOrDefault(recipe =>
                string.Equals(recipe.Id, item.RecipeId, StringComparison.Ordinal));
            if (recipe is null)
            {
                throw new PptxValidationException(
                    "asset_plan_recipe_not_found",
                    $"assetPlan[{index}].recipe_id was not found in the immutable Brand Profile.");
            }

            if (!string.Equals(recipe.Purpose, item.Purpose, StringComparison.Ordinal)
                || !string.Equals(recipe.StyleDirectionId, styleDirection.Id, StringComparison.Ordinal))
            {
                throw new PptxValidationException(
                    "asset_plan_recipe_mismatch",
                    $"assetPlan[{index}] purpose and style direction must match the selected layout recipe.");
            }

            ValidateOptionalToken(item.CropIntent, SupportedCropIntents, $"assetPlan[{index}].crop_intent");
            ValidateOptionalToken(item.AspectRatio, SupportedAspectRatios, $"assetPlan[{index}].aspect_ratio");
            ValidateOptionalToken(item.TextSafeArea, SupportedTextSafeAreas, $"assetPlan[{index}].text_safe_area");
            ValidateOptionalIdentifier(item.AttributionRef, $"assetPlan[{index}].attribution_ref");
            ValidateOptionalAssetId(item.AssetId, $"assetPlan[{index}].asset_id");
            ValidateVisualObjectAssetIds(item.VisualObjectAssetIds, $"assetPlan[{index}].visual_object_asset_ids");
            ValidateOptionalIdentifier(
                item.ApprovedAssetCollectionId,
                $"assetPlan[{index}].approved_asset_collection_id");

            ValidateAcquisition(brief.SourcePolicy, item, recipe, profile, index);
        }

        return result;
    }

    private static void ValidateAcquisition(
        DesignSourcePolicy sourcePolicy,
        AssetPlanItem item,
        BrandLayoutRecipe recipe,
        BrandProfileSnapshot profile,
        int index)
    {
        var path = $"assetPlan[{index}]";
        switch (item.Acquisition)
        {
            case AssetAcquisition.NativeDraw:
                EnsureRecipeDoesNotRequireUnavailableAsset(recipe, path);
                if (item.Status != AssetPlanStatus.Ready
                    || item.LicenseStatus != AssetLicenseStatus.NotRequired
                    || item.Fallback is not (AssetFallback.None or AssetFallback.NativeDraw)
                    || item.AssetId is not null
                    || item.ApprovedAssetCollectionId is not null)
                {
                    throw new PptxValidationException(
                        "asset_plan_native_draw_invalid",
                        $"{path} nativeDraw must be ready, use notRequired license status, and have no approved library collection.");
                }

                break;
            case AssetAcquisition.None:
                if (item.Status != AssetPlanStatus.Omitted
                    || item.PreferredMedium != AssetPreferredMedium.None
                    || item.LicenseStatus != AssetLicenseStatus.NotRequired
                    || item.Fallback != AssetFallback.None
                    || item.AssetId is not null
                    || item.ApprovedAssetCollectionId is not null
                    || item.AttributionRef is not null
                    || item.CropIntent is not null
                    || item.AspectRatio is not null
                    || item.TextSafeArea is not null
                    || item.VisualObjectAssetIds is { Count: > 0 }
                    || recipe.RequiredAssetRoles.Count > 0)
                {
                    throw new PptxValidationException(
                        "asset_plan_omission_invalid",
                        $"{path} acquisition=none is valid only with preferred_medium=none, fallback=none, status=omitted, license_status=notRequired, no asset metadata fields, and a recipe whose required_asset_roles is empty. noAssetLayout is not valid with acquisition=none.");
                }

                break;
            case AssetAcquisition.UserUpload:
                if (sourcePolicy != DesignSourcePolicy.ApprovedOrUserProvided)
                {
                    throw new PptxValidationException(
                        "asset_plan_source_policy_mismatch",
                        $"{path} userUpload is not allowed by brief.source_policy.");
                }

                if (item.Status == AssetPlanStatus.Ready)
                {
                    if (item.AssetId is null
                        || item.Fallback != AssetFallback.None
                        || item.ApprovedAssetCollectionId is not null
                        || item.LicenseStatus != AssetLicenseStatus.UserProvided
                        || item.PreferredMedium is not (AssetPreferredMedium.Photo or AssetPreferredMedium.Illustration)
                        || item.CropIntent is null
                        || !string.Equals(item.TextSafeArea, "left", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(item.TextSafeArea, "right", StringComparison.OrdinalIgnoreCase)
                        || recipe.RequiredAssetRoles.Count == 0)
                    {
                        throw new PptxValidationException(
                            "asset_plan_user_upload_invalid",
                            $"{path} ready userUpload must use a registered asset_id, photo or illustration medium, fallback=none, userProvided license status, a crop_intent, text_safe_area=left|right, no approved collection, and a recipe with a required asset role.");
                    }
                }
                else
                {
                    ValidateImageFallback(item, recipe, path);
                }

                if (item.ApprovedAssetCollectionId is not null
                    || item.VisualObjectAssetIds is { Count: > 0 }
                    || item.LicenseStatus is not (AssetLicenseStatus.UserProvided or AssetLicenseStatus.Unknown))
                {
                    throw new PptxValidationException(
                        "asset_plan_user_upload_invalid",
                        $"{path} userUpload must not name an approved collection and must use userProvided or unknown license status.");
                }

                break;
            case AssetAcquisition.ApprovedLibrary:
                ValidateImageFallback(item, recipe, path);
                if (item.AssetId is not null)
                {
                    throw new PptxValidationException(
                        "asset_plan_approved_library_invalid",
                        $"{path} approvedLibrary cannot use a user-uploaded asset_id.");
                }
                if (item.VisualObjectAssetIds is { Count: > 0 })
                {
                    throw new PptxValidationException(
                        "asset_plan_visual_objects_invalid",
                        $"{path} approvedLibrary cannot bind native visual object assets.");
                }
                if (item.ApprovedAssetCollectionId is null
                    || !profile.Detail.ApprovedAssetCollectionIds.Contains(
                        item.ApprovedAssetCollectionId,
                        StringComparer.Ordinal)
                    || item.LicenseStatus is not (AssetLicenseStatus.Approved or AssetLicenseStatus.Unknown or AssetLicenseStatus.Restricted))
                {
                    throw new PptxValidationException(
                        "asset_plan_approved_library_invalid",
                        $"{path} approvedLibrary must use an exact collection ID from the profile and a recorded approved, unknown, or restricted license status.");
                }

                break;
            default:
                throw new PptxValidationException(
                    "asset_plan_acquisition_invalid",
                    $"{path}.acquisition is unsupported.");
        }
    }

    private static void ValidateImageFallback(AssetPlanItem item, BrandLayoutRecipe recipe, string path)
    {
        if (item.Status != AssetPlanStatus.FallbackSelected
            || item.Fallback is not (AssetFallback.NativeDraw or AssetFallback.NoAssetLayout)
            || item.AssetId is not null)
        {
            throw new PptxValidationException(
                "asset_plan_image_insertion_unavailable",
                $"{path} without a usable registered image must select status=fallbackSelected, fallback=nativeDraw or noAssetLayout, and omit asset_id before generation.");
        }

        EnsureRecipeDoesNotRequireUnavailableAsset(recipe, path);

        if (item.LicenseStatus is AssetLicenseStatus.Unknown or AssetLicenseStatus.Restricted
            && item.AttributionRef is not null)
        {
            throw new PptxValidationException(
                "asset_plan_unverified_attribution_invalid",
                $"{path} must not attach attribution to an unknown or restricted asset that will not be used.");
        }
    }

    private static void ValidateOptionalAssetId(string? value, string path)
    {
        if (value is not null && !ImageAssetIdRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "asset_plan_asset_id_invalid",
                $"{path} must be the lowercase opaque asset_id returned by pptx_register_uploaded_image_asset.");
        }
    }

    private static void ValidateVisualObjectAssetIds(IReadOnlyList<string>? values, string path)
    {
        if (values is null)
        {
            return;
        }

        if (values.Count is < 1 or > VisualObjectAssetRepository.MaximumObjectsPerSlide
            || values.Any(static value => !ImageAssetIdRegex().IsMatch(value))
            || values.Distinct(StringComparer.Ordinal).Count() != values.Count)
        {
            throw new PptxValidationException(
                "asset_plan_visual_objects_invalid",
                $"{path} must contain 1-{VisualObjectAssetRepository.MaximumObjectsPerSlide} unique lowercase opaque IDs returned by pptx_prepare_visual_objects.");
        }
    }

    private static void EnsureRecipeDoesNotRequireUnavailableAsset(
        BrandLayoutRecipe recipe,
        string path)
    {
        if (recipe.RequiredAssetRoles.Count > 0)
        {
            throw new PptxValidationException(
                "asset_plan_recipe_requires_unavailable_asset",
                $"{path}.recipe_id requires an external asset, but phase 1 cannot insert images. Select a fallback recipe without required asset roles.");
        }
    }

    private void EnsureCapacityUnsafe()
    {
        var pendingCount = preparedSelections.Values.Count(static selection => selection.SelectedOptionId is null);
        if (briefs.Count + pendingCount >= MaximumActiveBriefs)
        {
            throw new PptxValidationException(
                "design_brief_capacity_reached",
                "Too many validated or pending Design Briefs are active. Retry after an existing brief expires.");
        }
    }

    private void EnsurePerUserCapacityUnsafe(CallerContext caller)
    {
        var count = preparedSelections.Values.Count(selection =>
            !selection.StartCompleted
            && ownership.TryGetValue(selection.ChoiceSessionId, out var owner)
            && string.Equals(owner.UserScope, caller.UserScope, StringComparison.Ordinal));
        if (count >= MaximumPendingChoicesPerUser)
        {
            throw new PptxValidationException(
                "design_brief_user_capacity_reached",
                "This user has too many unresolved Design Brief cards. Complete or let an existing choice expire before preparing another.");
        }
    }

    private PreparedDesignBriefSelection? FindBlockingSelectionUnsafe(CallerContext caller) =>
        preparedSelections.Values
            .Where(static selection => !selection.StartCompleted)
            .Where(selection => IsOwnedByCallerUnsafe(selection.ChoiceSessionId, caller))
            .OrderByDescending(static selection => selection.ExpiresAt)
            .FirstOrDefault();

    private bool IsOwnedByCallerUnsafe(string id, CallerContext caller) =>
        ownership.TryGetValue(id, out var owner)
        && string.Equals(owner.UserScope, caller.UserScope, StringComparison.Ordinal)
        && string.Equals(owner.ConversationScope, caller.ConversationScope, StringComparison.Ordinal);

    private bool HasStartReservationForCallerUnsafe(CallerContext caller) =>
        conversationStartReservations.Contains(
            new BriefOwnership(caller.UserScope, caller.ConversationScope));

    private void InvalidateUnstartedBriefsUnsafe(CallerContext caller)
    {
        foreach (var pair in briefs)
        {
            if (!startedBriefs.ContainsKey(pair.Key)
                && !startReservations.Contains(pair.Key)
                && IsOwnedByCallerUnsafe(pair.Key, caller))
            {
                briefs.TryRemove(pair.Key, out _);
                ownership.TryRemove(pair.Key, out _);
            }
        }
    }

    private void PruneExpiredUnsafe()
    {
        var now = timeProvider.GetUtcNow();
        foreach (var pair in briefs)
        {
            if (pair.Value.ExpiresAt <= now && !startReservations.Contains(pair.Key))
            {
                briefs.TryRemove(pair.Key, out _);
                ownership.TryRemove(pair.Key, out _);
                startedBriefs.TryRemove(pair.Key, out _);
            }
        }

        foreach (var pair in preparedSelections)
        {
            if (pair.Value.ExpiresAt <= now && !pair.Value.StartReserved)
            {
                preparedSelections.TryRemove(pair.Key, out _);
                ownership.TryRemove(pair.Key, out _);
            }
        }
    }

    private static string NewOpaqueId() => Guid.NewGuid().ToString("N");

    private static void ValidateOpaqueId(string value, string field)
    {
        if (value is null || !OpaqueIdRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "design_brief_action_identifier_invalid",
                $"{field} must be an opaque server-issued identifier.");
        }
    }

    private static void ValidateOptionalToken(
        string? value,
        HashSet<string> supported,
        string path)
    {
        if (value is not null && !supported.Contains(value))
        {
            throw new PptxValidationException(
                "asset_plan_token_invalid",
                $"{path} contains an unsupported token.");
        }
    }

    private static void ValidateOptionalIdentifier(string? value, string path)
    {
        if (value is not null)
        {
            ValidateIdentifier(value, path);
        }
    }

    private static void ValidateIdentifier(string value, string path)
    {
        if (value is null || !OpaqueIdentifierRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "design_brief_identifier_invalid",
                $"{path} must be an opaque ASCII identifier, not a URL or path.");
        }
    }

    private static void ValidateText(string value, string path, int minimumLength, int maximumLength)
    {
        if (value is null
            || value.Length < minimumLength
            || value.Length > maximumLength
            || value.Any(character => char.IsControl(character) && character is not ('\n' or '\r' or '\t'))
            || value.Contains("://", StringComparison.OrdinalIgnoreCase)
            || FileSchemeRegex().IsMatch(value)
            || value.Contains("../", StringComparison.Ordinal)
            || value.Contains("..\\", StringComparison.Ordinal)
            || value.StartsWith('/')
            || value.StartsWith("\\\\", StringComparison.Ordinal)
            || AbsolutePosixPathRegex().IsMatch(value)
            || UncPathRegex().IsMatch(value)
            || WindowsPathRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "design_brief_text_invalid",
                $"{path} must contain {minimumLength} to {maximumLength} characters and no URL, path, or control character.");
        }
    }

    [GeneratedRegex("\\A[0-9a-fA-F]{32}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex OpaqueIdRegex();

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex OpaqueIdentifierRegex();

    [GeneratedRegex("\\A[0-9a-f]{32}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex ImageAssetIdRegex();

    [GeneratedRegex("(?<![A-Za-z0-9_])[A-Za-z]:[\\\\/]", RegexOptions.CultureInvariant)]
    private static partial Regex WindowsPathRegex();

    [GeneratedRegex("(?i:(?<![A-Za-z0-9_])file:)", RegexOptions.CultureInvariant)]
    private static partial Regex FileSchemeRegex();

    [GeneratedRegex("""(?<![A-Za-z0-9_])/[A-Za-z0-9._-]+(?:/[^\s"')\]}>,;:]+)*""", RegexOptions.CultureInvariant)]
    private static partial Regex AbsolutePosixPathRegex();

    [GeneratedRegex("""(?<![A-Za-z0-9_])\\\\[^\\\s"')\]}>,;:]+\\[^\s"')\]}>,;:]+""", RegexOptions.CultureInvariant)]
    private static partial Regex UncPathRegex();

    private sealed record BriefOwnership(string UserScope, string ConversationScope);

    private sealed record PreparedDesignBriefSelection(
        string ChoiceSessionId,
        string BriefId,
        DateTimeOffset ExpiresAt,
        string ProfileId,
        string ProfileVersion,
        string ProfileContentHash,
        IReadOnlyDictionary<string, PreparedDesignBriefOption> Options,
        string? SelectedOptionId,
        bool StartReserved,
        bool StartCompleted);
}

internal enum DesignBriefCardOptionKind
{
    Recommended,
    Alternative,
    NoPhoto,
}

internal sealed record DesignBriefCardAssetSummary(
    int ReadyCount,
    int NeedsUserCount,
    int FallbackSelectedCount,
    int OmittedCount);

internal sealed record PreparedDesignBriefOption(
    string OptionId,
    DesignBriefCardOptionKind Kind,
    [property: System.Text.Json.Serialization.JsonIgnore] ValidatedDesignBriefBinding Binding,
    DesignBriefCardAssetSummary AssetSummary);

internal sealed record PreparedDesignBriefCard(
    string ChoiceSessionId,
    DateTimeOffset ExpiresAt,
    PreparedDesignBriefOption Recommended,
    IReadOnlyList<PreparedDesignBriefOption> Alternatives,
    PreparedDesignBriefOption? NoPhoto);
