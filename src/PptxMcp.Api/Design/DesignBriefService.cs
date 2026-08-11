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
    private readonly ConcurrentDictionary<string, BriefOwnership> ownership = new(StringComparer.Ordinal);

    public DesignBriefValidationView Validate(
        CallerContext caller,
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan)
    {
        ArgumentNullException.ThrowIfNull(caller);
        ArgumentNullException.ThrowIfNull(brief);
        ArgumentNullException.ThrowIfNull(assetPlan);
        PruneExpired();
        if (briefs.Count >= MaximumActiveBriefs)
        {
            throw new PptxValidationException(
                "design_brief_capacity_reached",
                "Too many validated Design Briefs are active. Retry after an existing brief expires.");
        }

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
        var now = timeProvider.GetUtcNow();
        var expiresAt = now.AddMinutes(options.DesignBriefLifetimeMinutes);
        while (true)
        {
            var id = Guid.NewGuid().ToString("N");
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
            var binding = new ValidatedDesignBriefBinding(
                id,
                brief,
                profile,
                styleDirection,
                theme,
                design,
                planBySlide,
                expiresAt);
            if (briefs.TryAdd(id, binding)
                && ownership.TryAdd(id, new BriefOwnership(caller.UserScope, caller.ConversationScope)))
            {
                var summary = new AssetPlanSummary(
                    assetPlan.Count(item => item.Acquisition == AssetAcquisition.NativeDraw),
                    assetPlan.Count(item => item.Acquisition == AssetAcquisition.None),
                    assetPlan.Count(item => item.Status == AssetPlanStatus.FallbackSelected));
                return new DesignBriefValidationView(
                    id,
                    "validated",
                    expiresAt,
                    brief.BrandProfile,
                    brief.StyleDirectionId,
                    brief.ExpectedSlideCount,
                    summary,
                    "Call pptx_start_visual_deck with this brief_id, the same expectedSlideCount, and the profile template_source. Omit theme and design because the validated profile direction supplies them.");
            }

            briefs.TryRemove(id, out _);
            ownership.TryRemove(id, out _);
        }
    }

    public ValidatedDesignBriefBinding? AuthorizeForStart(
        CallerContext caller,
        string? briefId,
        int expectedSlideCount,
        string templateSourceFileId,
        VisualThemeSpec? requestedTheme,
        VisualDesignSpec? requestedDesign)
    {
        ArgumentNullException.ThrowIfNull(caller);
        if (string.IsNullOrWhiteSpace(briefId))
        {
            if (options.RequireDesignBrief)
            {
                throw new PptxValidationException(
                    "design_brief_required",
                    "This deployment requires a validated Design Brief. Call pptx_get_design_catalog and pptx_validate_design_brief before pptx_start_visual_deck.");
            }

            return null;
        }

        if (!OpaqueIdRegex().IsMatch(briefId)
            || !briefs.TryGetValue(briefId.ToLowerInvariant(), out var brief)
            || !ownership.TryGetValue(briefId.ToLowerInvariant(), out var owner)
            || !string.Equals(owner.UserScope, caller.UserScope, StringComparison.Ordinal)
            || !string.Equals(owner.ConversationScope, caller.ConversationScope, StringComparison.Ordinal))
        {
            throw new PptxValidationException(
                "design_brief_not_found",
                "The Design Brief was not found for this user and conversation. Validate a new brief in this conversation.");
        }

        if (brief.ExpiresAt <= timeProvider.GetUtcNow())
        {
            briefs.TryRemove(brief.BriefId, out _);
            ownership.TryRemove(brief.BriefId, out _);
            throw new PptxValidationException(
                "design_brief_expired",
                "The Design Brief expired. Refresh the design catalog and validate the brief again.");
        }

        if (brief.Brief.ExpectedSlideCount != expectedSlideCount)
        {
            throw new PptxValidationException(
                "design_brief_slide_count_mismatch",
                $"expectedSlideCount must remain {brief.Brief.ExpectedSlideCount}, as finalized in the Design Brief.");
        }

        if (!string.Equals(brief.Profile.TemplateSource, templateSourceFileId.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            throw new PptxValidationException(
                "design_brief_template_mismatch",
                $"templateSourceFileId must remain {brief.Profile.TemplateSource}, as fixed by the Brand Profile.");
        }

        if (requestedTheme is not null || requestedDesign is not null)
        {
            throw new PptxValidationException(
                "design_brief_creative_direction_conflict",
                "Omit theme and design when using briefId; the validated Brand Profile version and style direction supply immutable values.");
        }

        return brief;
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
                    || item.ApprovedAssetCollectionId is not null
                    || item.AttributionRef is not null
                    || item.CropIntent is not null
                    || item.AspectRatio is not null
                    || item.TextSafeArea is not null
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

                ValidatePhaseOneFallback(item, path);
                EnsureRecipeDoesNotRequireUnavailableAsset(recipe, path);
                if (item.ApprovedAssetCollectionId is not null
                    || item.LicenseStatus is not (AssetLicenseStatus.UserProvided or AssetLicenseStatus.Unknown))
                {
                    throw new PptxValidationException(
                        "asset_plan_user_upload_invalid",
                        $"{path} userUpload must not name an approved collection and must use userProvided or unknown license status.");
                }

                break;
            case AssetAcquisition.ApprovedLibrary:
                ValidatePhaseOneFallback(item, path);
                EnsureRecipeDoesNotRequireUnavailableAsset(recipe, path);
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

    private static void ValidatePhaseOneFallback(AssetPlanItem item, string path)
    {
        if (item.Status != AssetPlanStatus.FallbackSelected
            || item.Fallback is not (AssetFallback.NativeDraw or AssetFallback.NoAssetLayout))
        {
            throw new PptxValidationException(
                "asset_plan_image_insertion_unavailable",
                $"{path} plans a non-native image source, but phase 1 cannot insert images. Select status=fallbackSelected and fallback=nativeDraw or noAssetLayout before generation.");
        }

        if (item.LicenseStatus is AssetLicenseStatus.Unknown or AssetLicenseStatus.Restricted
            && item.AttributionRef is not null)
        {
            throw new PptxValidationException(
                "asset_plan_unverified_attribution_invalid",
                $"{path} must not attach attribution to an unknown or restricted asset that will not be used.");
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

    private void PruneExpired()
    {
        var now = timeProvider.GetUtcNow();
        foreach (var pair in briefs)
        {
            if (pair.Value.ExpiresAt <= now)
            {
                briefs.TryRemove(pair.Key, out _);
                ownership.TryRemove(pair.Key, out _);
            }
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

    [GeneratedRegex("(?<![A-Za-z0-9_])[A-Za-z]:[\\\\/]", RegexOptions.CultureInvariant)]
    private static partial Regex WindowsPathRegex();

    [GeneratedRegex("(?i:(?<![A-Za-z0-9_])file:)", RegexOptions.CultureInvariant)]
    private static partial Regex FileSchemeRegex();

    [GeneratedRegex("""(?<![A-Za-z0-9_])/[A-Za-z0-9._-]+(?:/[^\s"')\]}>,;:]+)*""", RegexOptions.CultureInvariant)]
    private static partial Regex AbsolutePosixPathRegex();

    [GeneratedRegex("""(?<![A-Za-z0-9_])\\\\[^\\\s"')\]}>,;:]+\\[^\s"')\]}>,;:]+""", RegexOptions.CultureInvariant)]
    private static partial Regex UncPathRegex();

    private sealed record BriefOwnership(string UserScope, string ConversationScope);
}
