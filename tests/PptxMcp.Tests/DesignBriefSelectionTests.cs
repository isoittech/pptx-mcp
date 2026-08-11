using System.Text.Json.Nodes;
using System.Text.Json;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class DesignBriefSelectionTests
{
    private static readonly CallerContext Caller = new("user-a", "conversation-a", "message-a");

    [Fact]
    public void CardRequiresAtLeastTwoMateriallyDifferentChoices()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        try
        {
            var (service, catalog, _) = CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));

            var error = Assert.Throws<PptxValidationException>(() =>
                service.Prepare(Caller, brief, assetPlan, null, null));

            Assert.Equal("design_brief_card_choice_required", error.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void CardSelectionIsOwnerBoundIdempotentBeforeStartAndReplaySafe()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        AddAlternativeDirection(root, materiallyDifferent: true);
        try
        {
            var (service, catalog, _) = CreateService(root, requireDesignBrief: false);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var old = service.Validate(Caller, brief, assetPlan);
            var card = service.Prepare(
                Caller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null);

            Assert.Matches("^[0-9a-f]{32}$", card.ChoiceSessionId);
            Assert.Matches("^[0-9a-f]{32}$", card.Recommended.OptionId);
            Assert.Single(card.Alternatives);
            Assert.Throws<PptxValidationException>(() =>
                service.Validate(Caller, brief, assetPlan));
            Assert.Equal(
                "design_brief_not_confirmed",
                Assert.Throws<PptxValidationException>(() =>
                    service.AuthorizeForStart(Caller, old.BriefId, 2, "none", null, null)).Code);
            Assert.Equal(
                "design_brief_not_confirmed",
                Assert.Throws<PptxValidationException>(() =>
                    service.AuthorizeForStart(Caller, null, 2, "none", null, null)).Code);

            var otherCaller = new CallerContext("user-a", "conversation-b", "message-b");
            Assert.Equal(
                "design_brief_action_not_found",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    otherCaller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId)).Code);
            Assert.Equal(
                "design_brief_action_tampered",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    Guid.NewGuid().ToString("N"))).Code);

            var applied = service.ApplyCardSelection(
                Caller,
                card.ChoiceSessionId,
                card.Recommended.OptionId);
            var duplicate = service.ApplyCardSelection(
                Caller,
                card.ChoiceSessionId,
                card.Recommended.OptionId);

            Assert.Equal(applied.BriefId, duplicate.BriefId);
            Assert.Equal(card.Recommended.Binding.BriefId, applied.BriefId);
            Assert.Equal(
                "design_brief_action_replayed",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Alternatives[0].OptionId)).Code);
            Assert.Equal(
                "design_brief_selection_required",
                Assert.Throws<PptxValidationException>(() =>
                    service.AuthorizeForStart(Caller, null, 2, "none", null, null)).Code);
            var binding = service.AuthorizeForStart(
                Caller,
                applied.BriefId,
                2,
                "none",
                null,
                null)!;
            Assert.Equal(DesignBriefSelectionSource.UserCard, binding.SelectionSource);
            var drafts = new VisualDeckDraftService(
                Options.Create(new PptxMcpOptions { MaxSlides = 50 }),
                TimeProvider.System);
            var draft = drafts.Begin(
                Caller,
                "Selected brief audit",
                2,
                binding.Theme,
                null,
                "en-US",
                binding.Design,
                "none",
                "auto",
                designBrief: binding);
            drafts.AddSlides(
                Caller,
                draft.DraftId,
                null,
                [new VisualSlideSpec(
                    VisualSlideKind.Title,
                    "Decision",
                    Density: "airy",
                    RecipeId: "cover-airy")]);
            drafts.AddSlides(
                Caller,
                draft.DraftId,
                null,
                [new VisualSlideSpec(
                    VisualSlideKind.Metrics,
                    "Evidence",
                    Metrics:
                    [
                        new VisualMetricSpec("1", "First"),
                        new VisualMetricSpec("2", "Second"),
                        new VisualMetricSpec("3", "Third"),
                    ],
                    Variant: "spotlight",
                    RecipeId: "kpi-balanced")]);
            var deck = drafts.AcquireForSubmission(Caller, draft.DraftId).Deck!;
            Assert.Equal(
                DesignBriefSelectionSource.UserCard,
                deck.BrandProfileBinding?.DesignBriefAudit?.SelectionSource);
            var auditJson = JsonSerializer.Serialize(deck);
            Assert.Contains("\"selection_source\":\"userCard\"", auditJson, StringComparison.Ordinal);
            Assert.DoesNotContain("choiceSessionId", auditJson, StringComparison.Ordinal);
            Assert.DoesNotContain("optionId", auditJson, StringComparison.Ordinal);
            Assert.DoesNotContain(card.ChoiceSessionId, auditJson, StringComparison.Ordinal);

            Assert.True(service.ReserveStart(Caller, binding));
            Assert.Equal(
                "design_brief_start_in_progress",
                Assert.Throws<PptxValidationException>(() =>
                    service.Validate(Caller, brief, assetPlan)).Code);
            service.ReleaseStartReservation(Caller, binding);
            Assert.Equal(
                applied.BriefId,
                service.ApplyCardSelection(Caller, card.ChoiceSessionId, card.Recommended.OptionId).BriefId);
            Assert.Equal(
                "design_brief_choice_already_applied",
                Assert.Throws<PptxValidationException>(() => service.Prepare(
                    Caller,
                    brief,
                    assetPlan,
                    [new DesignBriefStyleAlternative(
                        "report",
                        "detailed",
                        ["cover-report-airy", "kpi-report-detailed"])],
                    null,
                    replacePendingChoice: true)).Code);

            Assert.True(service.ReserveStart(Caller, binding));
            service.MarkStartSucceeded(Caller, binding.BriefId);
            Assert.Equal(
                "design_brief_action_already_started",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId)).Code);
            Assert.Equal(
                "design_brief_already_started",
                Assert.Throws<PptxValidationException>(() => service.AuthorizeForStart(
                    Caller,
                    binding.BriefId,
                    2,
                    "none",
                    null,
                    null,
                    userRequestedNewWorkflow: true)).Code);

            var idempotent = service.AuthorizeForStart(
                Caller,
                binding.BriefId,
                2,
                "none",
                null,
                null)!;
            Assert.True(service.ReserveStart(Caller, idempotent));
            service.MarkStartSucceeded(Caller, idempotent.BriefId);
            var direct = service.Validate(Caller, brief, assetPlan);
            Assert.Equal(
                DesignBriefSelectionSource.AgentDefault,
                service.AuthorizeForStart(
                    Caller,
                    direct.BriefId,
                    2,
                    "none",
                    null,
                    null)?.SelectionSource);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void PendingCardCanBeExplicitlyReplacedButAppliedCardCannot()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        AddAlternativeDirection(root, materiallyDifferent: true);
        try
        {
            var (service, catalog, _) = CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var alternatives = new[]
            {
                new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"]),
            };
            var first = service.Prepare(Caller, brief, assetPlan, alternatives, null);
            Assert.Equal(
                "design_brief_choice_pending",
                Assert.Throws<PptxValidationException>(() =>
                    service.Prepare(Caller, brief, assetPlan, alternatives, null)).Code);

            var replacement = service.Prepare(
                Caller,
                brief,
                assetPlan,
                alternatives,
                null,
                replacePendingChoice: true);

            Assert.NotEqual(first.ChoiceSessionId, replacement.ChoiceSessionId);
            Assert.Equal(
                "design_brief_action_not_found",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    first.ChoiceSessionId,
                    first.Recommended.OptionId)).Code);
            Assert.NotNull(service.ApplyCardSelection(
                Caller,
                replacement.ChoiceSessionId,
                replacement.Recommended.OptionId));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void PendingCardCanBeCancelledWithoutIdentifiersButSelectedCardCannot()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        AddAlternativeDirection(root, materiallyDifferent: true);
        try
        {
            var (service, catalog, _) = CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            Assert.Equal(
                "design_brief_choice_not_pending",
                Assert.Throws<PptxValidationException>(() =>
                    service.CancelPendingSelection(Caller)).Code);

            var card = service.Prepare(
                Caller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null);
            var cancelled = service.CancelPendingSelection(Caller);

            Assert.Equal("cancelled", cancelled.Status);
            Assert.Equal(
                "design_brief_action_not_found",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId)).Code);
            var directAfterCancel = service.Validate(Caller, brief, assetPlan);
            Assert.Equal(
                DesignBriefSelectionSource.AgentDefault,
                service.AuthorizeForStart(
                    Caller,
                    directAfterCancel.BriefId,
                    2,
                    "none",
                    null,
                    null)?.SelectionSource);

            var selectedCaller = new CallerContext("selected-user", "selected-conversation", "selected-message");
            var selectedCard = service.Prepare(
                selectedCaller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null);
            service.ApplyCardSelection(
                selectedCaller,
                selectedCard.ChoiceSessionId,
                selectedCard.Recommended.OptionId);

            Assert.Equal(
                "design_brief_choice_cancel_forbidden",
                Assert.Throws<PptxValidationException>(() =>
                    service.CancelPendingSelection(selectedCaller)).Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ExpiredAndPrunedCardNeverActivatesAStartBrief()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        AddAlternativeDirection(root, materiallyDifferent: true);
        try
        {
            var (service, catalog, clock) = CreateService(root, lifetimeMinutes: 1);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var card = service.Prepare(
                Caller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null);
            clock.Advance(TimeSpan.FromMinutes(2));

            Assert.Equal(
                "design_brief_action_expired",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId)).Code);
            Assert.Equal(
                "design_brief_action_not_found",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId)).Code);
            Assert.Equal(
                "design_brief_not_found",
                Assert.Throws<PptxValidationException>(() => service.AuthorizeForStart(
                    Caller,
                    card.Recommended.Binding.BriefId,
                    2,
                    "none",
                    null,
                    null)).Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void EffectiveRenderFingerprintRejectsCosmeticIdsCaseOnlyTokensAndOverriddenPreset()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        AddAlternativeDirection(root, materiallyDifferent: false);
        try
        {
            var (service, catalog, _) = CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));

            var error = Assert.Throws<PptxValidationException>(() => service.Prepare(
                Caller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "balanced",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null));

            Assert.Equal("design_brief_alternative_has_no_visual_difference", error.Code);

            var densityOnly = Assert.Throws<PptxValidationException>(() => service.Prepare(
                new CallerContext("user-density", "conversation-density", "message-density"),
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null));
            Assert.Equal("design_brief_alternative_has_no_visual_difference", densityOnly.Code);

            var nullRecipes = Assert.Throws<PptxValidationException>(() => service.Prepare(
                new CallerContext("user-b", "conversation-b", "message-b"),
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative("report", "balanced", null!)],
                null));
            Assert.Equal("design_brief_alternatives_invalid", nullRecipes.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void PhotoFreeChoiceCanonicalizesAssetFieldsAndMustChangeEffectiveRecipe()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        AddNoPhotoRecipes(root);
        try
        {
            var (service, catalog, _) = CreateService(root);
            var (baseBrief, basePlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var brief = baseBrief with { SourcePolicy = DesignSourcePolicy.ApprovedOrUserProvided };
            var assetPlan = basePlan.ToArray();
            assetPlan[1] = assetPlan[1] with
            {
                PreferredMedium = AssetPreferredMedium.Photo,
                Acquisition = AssetAcquisition.UserUpload,
                Fallback = AssetFallback.NativeDraw,
                Status = AssetPlanStatus.FallbackSelected,
                LicenseStatus = AssetLicenseStatus.UserProvided,
                AttributionRef = "user-photo-record",
                CropIntent = "cover",
                AspectRatio = "landscape16x9",
                TextSafeArea = "left",
            };
            var card = service.Prepare(
                Caller,
                brief,
                assetPlan,
                null,
                [new DesignBriefNoPhotoOverride(
                    2,
                    "kpi-no-photo",
                    AssetPreferredMedium.NativeDiagram,
                    AssetAcquisition.NativeDraw)]);

            var photoFree = Assert.IsType<PreparedDesignBriefOption>(card.NoPhoto);
            var item = photoFree.Binding.AssetPlan[2];
            Assert.Equal(AssetFallback.None, item.Fallback);
            Assert.Equal(AssetPlanStatus.Ready, item.Status);
            Assert.Equal(AssetLicenseStatus.NotRequired, item.LicenseStatus);
            Assert.Null(item.ApprovedAssetCollectionId);
            Assert.Null(item.AttributionRef);
            Assert.Null(item.CropIntent);
            Assert.Null(item.AspectRatio);
            Assert.Null(item.TextSafeArea);

            var sameVisual = Assert.Throws<PptxValidationException>(() => service.Prepare(
                new CallerContext("user-b", "conversation-b", "message-b"),
                brief,
                assetPlan,
                null,
                [new DesignBriefNoPhotoOverride(
                    2,
                    "kpi-no-photo-same",
                    AssetPreferredMedium.NativeDiagram,
                    AssetAcquisition.NativeDraw)]));
            Assert.Equal("design_brief_no_photo_has_no_visual_difference", sameVisual.Code);

            var nullOverride = Assert.Throws<PptxValidationException>(() => service.Prepare(
                new CallerContext("user-c", "conversation-c", "message-c"),
                brief,
                assetPlan,
                null,
                [null!]));
            Assert.Equal("design_brief_no_photo_invalid", nullOverride.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void DiscardPreparedRollsBackCardResourceFailureState()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        AddAlternativeDirection(root, materiallyDifferent: true);
        try
        {
            var (service, catalog, _) = CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(GetReference(catalog));
            var card = service.Prepare(
                Caller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null);

            service.DiscardPrepared(Caller, card.ChoiceSessionId);

            Assert.NotNull(service.Validate(Caller, brief, assetPlan));
            Assert.Equal(
                "design_brief_action_not_found",
                Assert.Throws<PptxValidationException>(() => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId)).Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    internal static (DesignBriefService Service, BrandProfileCatalog Catalog, MutableTimeProvider Clock) CreateService(
        string root,
        bool requireDesignBrief = true,
        int lifetimeMinutes = 60)
    {
        var clock = new MutableTimeProvider(new DateTimeOffset(2026, 8, 11, 0, 0, 0, TimeSpan.Zero));
        var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief, lifetimeMinutes);
        var catalog = new BrandProfileCatalog(options);
        return (new DesignBriefService(options, clock, catalog), catalog, clock);
    }

    internal static BrandProfileReference GetReference(BrandProfileCatalog catalog)
    {
        var summary = catalog.Query("general").Profiles.Single().Summary;
        return new BrandProfileReference(summary.Id, summary.Version, summary.ContentHash);
    }

    internal static void AddAlternativeDirection(string root, bool materiallyDifferent)
    {
        var path = Path.Combine(root, "general", "brand-profile.json");
        var manifest = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        var styles = manifest["style_directions"]!.AsArray();
        var recipes = manifest["layout_recipes"]!.AsArray();
        var samples = manifest["samples"]!.AsArray();
        styles.Add(JsonNode.Parse(materiallyDifferent
            ? """{"id":"report","name":"Report","summary":"Dense editorial reporting.","recommended_for":["cover","kpi"],"design_style":"editorial","default_density":"detailed","supported_densities":["airy","balanced","detailed"],"motif":"ribbon","theme_preset":"ocean"}"""
            : """{"id":"report","name":"Report alias","summary":"An alias with no render difference.","recommended_for":["cover","kpi"],"design_style":"EXECUTIVE","default_density":"BALANCED","supported_densities":["AIRY","BALANCED","DETAILED"],"motif":"GEOMETRIC","theme_preset":"OCEAN"}"""));
        recipes.Add(JsonNode.Parse(materiallyDifferent
            ? """{"id":"cover-report-airy","purpose":"cover","semantic_kind":"Title","variant":"auto","density":"airy","style_direction_id":"report","required_asset_roles":[],"sample_ids":["sample-cover-report"]}"""
            : """{"id":"cover-report-airy","purpose":"cover","semantic_kind":"Title","variant":"AUTO","density":"AIRY","style_direction_id":"report","required_asset_roles":[],"sample_ids":["sample-cover-report"]}"""));
        recipes.Add(JsonNode.Parse(materiallyDifferent
            ? """{"id":"kpi-report-detailed","purpose":"kpi","semantic_kind":"Cards","variant":"spotlight","density":"detailed","style_direction_id":"report","required_asset_roles":[],"sample_ids":["sample-kpi-report"]}"""
            : """{"id":"kpi-report-detailed","purpose":"kpi","semantic_kind":"Metrics","variant":"SPOTLIGHT","density":"BALANCED","style_direction_id":"report","required_asset_roles":[],"sample_ids":["sample-kpi-report"],"item_count":3}"""));
        samples.Add(JsonNode.Parse("""{"id":"sample-cover-report","title":"Report cover","summary":"Alternative cover sample.","purpose":"cover","density":"airy","style_direction_id":"report","recipe_id":"cover-report-airy","information_level":"low"}"""));
        samples.Add(JsonNode.Parse(materiallyDifferent
            ? """{"id":"sample-kpi-report","title":"Detailed report","summary":"Alternative detailed composition.","purpose":"kpi","density":"detailed","style_direction_id":"report","recipe_id":"kpi-report-detailed","information_level":"high"}"""
            : """{"id":"sample-kpi-report","title":"KPI alias","summary":"Same effective composition.","purpose":"kpi","density":"balanced","style_direction_id":"report","recipe_id":"kpi-report-detailed","information_level":"medium"}"""));
        File.WriteAllText(path, manifest.ToJsonString());
    }

    internal static void AddNoPhotoRecipes(string root)
    {
        var path = Path.Combine(root, "general", "brand-profile.json");
        var manifest = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        var recipes = manifest["layout_recipes"]!.AsArray();
        recipes.Add(JsonNode.Parse("""{"id":"kpi-no-photo","purpose":"kpi","semantic_kind":"Cards","variant":"spotlight","density":"balanced","style_direction_id":"standard","required_asset_roles":[],"sample_ids":[]}"""));
        recipes.Add(JsonNode.Parse("""{"id":"kpi-no-photo-same","purpose":"kpi","semantic_kind":"Metrics","variant":"spotlight","density":"balanced","style_direction_id":"standard","required_asset_roles":[],"sample_ids":[],"item_count":3}"""));
        File.WriteAllText(path, manifest.ToJsonString());
    }

    internal sealed class MutableTimeProvider(DateTimeOffset now) : TimeProvider
    {
        private DateTimeOffset now = now;

        public override DateTimeOffset GetUtcNow() => now;

        public void Advance(TimeSpan duration) => now = now.Add(duration);
    }
}
