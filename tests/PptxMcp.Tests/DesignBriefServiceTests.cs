using Microsoft.AspNetCore.Http;
using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Security;
using PptxMcp.Storage;
using PptxMcp.Tools;

namespace PptxMcp.Tests;

public sealed class DesignBriefServiceTests
{
    private static readonly CallerContext Caller = new("user-a", "conversation-a", "message-a");

    [Fact]
    public void IssuesOpaqueBriefBoundToCallerConversationProfileAndCreativeDirection()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var clock = new MutableTimeProvider(new DateTimeOffset(2026, 8, 11, 0, 0, 0, TimeSpan.Zero));
            var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, clock, catalog);
            var reference = GetReference(catalog);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(reference);

            var validated = service.Validate(Caller, brief, assetPlan);
            var binding = service.AuthorizeForStart(
                Caller,
                validated.BriefId,
                expectedSlideCount: 2,
                templateSourceFileId: "none",
                requestedTheme: null,
                requestedDesign: null);

            Assert.Matches("^[0-9a-f]{32}$", validated.BriefId);
            Assert.Equal("validated", validated.Status);
            Assert.NotNull(binding);
            Assert.Equal(reference.ContentHash, binding.Profile.ContentHash);
            Assert.Equal("minimal", binding.Theme.Preset);
            Assert.Equal("executive", binding.Design.Style);
            Assert.Equal("balanced", binding.Design.Density);
            Assert.Null(binding.Theme.FontFace);
            Assert.Equal("Aptos", binding.Theme.HeadingFontFace);
            Assert.Equal("Aptos", binding.Theme.BodyFontFace);
            Assert.Equal("#FFFFFF", binding.Theme.SurfaceColor);
            Assert.Equal("#66717C", binding.Theme.MutedTextColor);
            Assert.Equal("#2B7A4B", binding.Theme.PositiveColor);
            Assert.Equal(3, binding.Theme.DataSeriesColors?.Count);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsBriefFromAnotherConversationAndAfterExpiry()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var clock = new MutableTimeProvider(new DateTimeOffset(2026, 8, 11, 0, 0, 0, TimeSpan.Zero));
            var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true, lifetimeMinutes: 60);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, clock, catalog);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var validated = service.Validate(Caller, brief, assetPlan);
            var otherConversation = new CallerContext("user-a", "conversation-b", "message-b");

            var ownershipError = Assert.Throws<PptxValidationException>(() => service.AuthorizeForStart(
                otherConversation,
                validated.BriefId,
                2,
                "none",
                null,
                null));
            clock.Advance(TimeSpan.FromMinutes(61));
            var expiryError = Assert.Throws<PptxValidationException>(() => service.AuthorizeForStart(
                Caller,
                validated.BriefId,
                2,
                "none",
                null,
                null));

            Assert.Equal("design_brief_not_found", ownershipError.Code);
            Assert.Equal("design_brief_expired", expiryError.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void GateIsBackwardCompatibleWhenDisabledAndRequiredWhenEnabled()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var optionalOptions = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: false);
            var requiredOptions = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true);
            var optional = new DesignBriefService(
                optionalOptions,
                TimeProvider.System,
                new BrandProfileCatalog(optionalOptions));
            var required = new DesignBriefService(
                requiredOptions,
                TimeProvider.System,
                new BrandProfileCatalog(requiredOptions));

            var noBinding = optional.AuthorizeForStart(Caller, null, 2, "none", null, null);
            var error = Assert.Throws<PptxValidationException>(() =>
                required.AuthorizeForStart(Caller, null, 2, "none", null, null));

            Assert.Null(noBinding);
            Assert.Equal("design_brief_required", error.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsUnresolvedQuestionsAndNonNativeAssetsWithoutSelectedFallback()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var options = BrandProfileTestFactory.CreateOptions(root);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, TimeProvider.System, catalog);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var unresolved = brief with { QuestionsForUser = ["Is an approved image available?"] };
            var userUploadBrief = brief with { SourcePolicy = DesignSourcePolicy.ApprovedOrUserProvided };
            var pendingAssetPlan = assetPlan.ToArray();
            pendingAssetPlan[1] = pendingAssetPlan[1] with
            {
                PreferredMedium = AssetPreferredMedium.Photo,
                Acquisition = AssetAcquisition.UserUpload,
                Fallback = AssetFallback.AskUser,
                Status = AssetPlanStatus.NeedsUser,
                LicenseStatus = AssetLicenseStatus.Unknown,
            };

            var questionError = Assert.Throws<PptxValidationException>(() =>
                service.Validate(Caller, unresolved, assetPlan));
            var assetError = Assert.Throws<PptxValidationException>(() =>
                service.Validate(Caller, userUploadBrief, pendingAssetPlan));

            Assert.Equal("design_brief_questions_unresolved", questionError.Code);
            Assert.Equal("asset_plan_image_insertion_unavailable", assetError.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void OmittedAssetReturnsCanonicalOneRetryCorrection()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var options = BrandProfileTestFactory.CreateOptions(root);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, TimeProvider.System, catalog);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var invalidFallbackPlan = assetPlan.ToArray();
            invalidFallbackPlan[0] = invalidFallbackPlan[0] with
            {
                Fallback = AssetFallback.NoAssetLayout,
            };
            var httpContext = new DefaultHttpContext();
            httpContext.Request.Headers["X-LibreChat-User-ID"] = "user-a";
            httpContext.Request.Headers["X-LibreChat-Conversation-ID"] = "conversation-a";
            var caller = new CallerContextAccessor(new HttpContextAccessor { HttpContext = httpContext });

            var result = PowerPointTools.ValidateDesignBrief(caller, service, brief, invalidFallbackPlan);

            var error = Assert.IsType<ToolValidationError>(result);
            Assert.Equal("asset_plan_omission_invalid", error.Code);
            Assert.Contains("fallback=none", error.Message, StringComparison.Ordinal);
            Assert.Contains("license_status=notRequired", error.Message, StringComparison.Ordinal);
            Assert.Contains("Never pair acquisition=none with noAssetLayout", error.Instruction, StringComparison.Ordinal);
            Assert.Contains("omit approved_asset_collection_id, attribution_ref, crop_intent, aspect_ratio, and text_safe_area", error.Instruction, StringComparison.Ordinal);

            var invalidMetadataPlan = assetPlan.ToArray();
            invalidMetadataPlan[0] = invalidMetadataPlan[0] with { AttributionRef = "unused-asset" };
            var metadataError = Assert.Throws<PptxValidationException>(() =>
                service.Validate(Caller, brief, invalidMetadataPlan));
            Assert.Equal("asset_plan_omission_invalid", metadataError.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void StartBindingRejectsChangedSlideCountTemplateAndExplicitCreativeOverrides()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, TimeProvider.System, catalog);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var validated = service.Validate(Caller, brief, assetPlan);

            var countError = Assert.Throws<PptxValidationException>(() => service.AuthorizeForStart(
                Caller, validated.BriefId, 3, "none", null, null));
            var templateError = Assert.Throws<PptxValidationException>(() => service.AuthorizeForStart(
                Caller, validated.BriefId, 2, "default", null, null));
            var directionError = Assert.Throws<PptxValidationException>(() => service.AuthorizeForStart(
                Caller,
                validated.BriefId,
                2,
                "none",
                new VisualThemeSpec("minimal"),
                null));

            Assert.Equal("design_brief_slide_count_mismatch", countError.Code);
            Assert.Equal("design_brief_template_mismatch", templateError.Code);
            Assert.Equal("design_brief_creative_direction_conflict", directionError.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void NativeDrawCannotSatisfyRecipeThatRequiresAnExternalAssetRole()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root, kpiRequiredAssetRoles: ["hero_photo"]);

        try
        {
            var options = BrandProfileTestFactory.CreateOptions(root);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, TimeProvider.System, catalog);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));

            var error = Assert.Throws<PptxValidationException>(() =>
                service.Validate(Caller, brief, assetPlan));

            Assert.Equal("asset_plan_recipe_requires_unavailable_asset", error.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData("Use the asset at (/mnt/brand/approved.png).")]
    [InlineData("Use the asset at /logo.png.")]
    [InlineData("Do not reveal /secret.")]
    [InlineData("Use the asset at \\\\server\\share\\approved.png.")]
    [InlineData("Use the asset at file:/mnt/brand/approved.png.")]
    public void RejectsEmbeddedAbsoluteFileLocatorsInBriefText(string assumptionText)
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var options = BrandProfileTestFactory.CreateOptions(root);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, TimeProvider.System, catalog);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var unsafeBrief = brief with
            {
                Assumptions = [new DesignAssumption(assumptionText, DesignAssumptionStatus.Inferred)],
            };

            var error = Assert.Throws<PptxValidationException>(() =>
                service.Validate(Caller, unsafeBrief, assetPlan));

            Assert.Equal("design_brief_text_invalid", error.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private static BrandProfileReference GetReference(BrandProfileCatalog catalog)
    {
        var summary = catalog.Query("general").Profiles.Single().Summary;
        return new BrandProfileReference(summary.Id, summary.Version, summary.ContentHash);
    }

    private sealed class MutableTimeProvider(DateTimeOffset now) : TimeProvider
    {
        private DateTimeOffset now = now;

        public override DateTimeOffset GetUtcNow() => now;

        public void Advance(TimeSpan duration) => now = now.Add(duration);
    }
}
