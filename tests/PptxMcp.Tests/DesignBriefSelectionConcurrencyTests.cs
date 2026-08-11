using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class DesignBriefSelectionConcurrencyTests
{
    private static readonly CallerContext Caller = new("phase2-user", "phase2-conversation", "phase2-message");

    [Fact]
    public void PendingCardSupersedesDirectBriefAndBlocksOptionalUnboundStart()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, _) = DesignBriefSelectionTests.CreateService(
                root,
                requireDesignBrief: false);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var direct = service.Validate(Caller, brief, assetPlan);
            var card = PrepareCard(service, brief, assetPlan);

            AssertCode(
                "design_brief_choice_pending",
                () => service.Validate(Caller, brief, assetPlan));
            AssertCode(
                "design_brief_choice_pending",
                () => PrepareCard(service, brief, assetPlan));
            AssertCode(
                "design_brief_not_confirmed",
                () => service.AuthorizeForStart(Caller, null, 2, "none", null, null));
            AssertCode(
                "design_brief_not_confirmed",
                () => service.AuthorizeForStart(Caller, direct.BriefId, 2, "none", null, null));
            AssertCode(
                "design_brief_not_confirmed",
                () => service.ReserveStart(Caller, null));

            var selected = service.ApplyCardSelection(
                Caller,
                card.ChoiceSessionId,
                card.Recommended.OptionId);

            AssertCode(
                "design_brief_selection_required",
                () => service.AuthorizeForStart(Caller, null, 2, "none", null, null));
            AssertCode(
                "design_brief_selection_superseded",
                () => service.AuthorizeForStart(Caller, direct.BriefId, 2, "none", null, null));
            Assert.Equal(
                selected.BriefId,
                service.AuthorizeForStart(Caller, selected.BriefId, 2, "none", null, null)?.BriefId);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ConcurrentApplyIsOwnerBoundAndSelectsExactlyOneServerOption()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, _) = DesignBriefSelectionTests.CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var card = PrepareCard(service, brief, assetPlan);

            AssertCode(
                "design_brief_action_not_found",
                () => service.ApplyCardSelection(
                    new CallerContext("other-user", "phase2-conversation", "other-message"),
                    card.ChoiceSessionId,
                    card.Recommended.OptionId));
            AssertCode(
                "design_brief_action_not_found",
                () => service.ApplyCardSelection(
                    new CallerContext("phase2-user", "other-conversation", "other-message"),
                    card.ChoiceSessionId,
                    card.Recommended.OptionId));

            using var start = new ManualResetEventSlim(false);
            var options = new[] { card.Recommended, card.Alternatives.Single() };
            var tasks = options.Select(option => Task.Run(() =>
            {
                start.Wait();
                try
                {
                    return new ApplyAttempt(
                        option,
                        service.ApplyCardSelection(Caller, card.ChoiceSessionId, option.OptionId),
                        null);
                }
                catch (PptxValidationException exception)
                {
                    return new ApplyAttempt(option, null, exception.Code);
                }
            })).ToArray();

            start.Set();
            var attempts = await Task.WhenAll(tasks);
            var winner = Assert.Single(attempts, static attempt => attempt.View is not null);
            var rejected = Assert.Single(attempts, static attempt => attempt.ErrorCode is not null);

            Assert.Equal("design_brief_action_replayed", rejected.ErrorCode);
            Assert.Equal(winner.Option.Binding.Brief.StyleDirectionId, winner.View!.StyleDirectionId);
            Assert.Equal(card.Recommended.Binding.BriefId, winner.View.BriefId);
            var binding = service.AuthorizeForStart(
                Caller,
                winner.View.BriefId,
                2,
                "none",
                null,
                null)!;
            Assert.Equal(winner.View.StyleDirectionId, binding.StyleDirection.Id);
            AssertCode(
                "design_brief_selection_superseded",
                () => service.AuthorizeForStart(
                    Caller,
                    Guid.NewGuid().ToString("N"),
                    2,
                    "none",
                    null,
                    null));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ConcurrentApplyAndCancelHasOneAtomicWinner()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, _) = DesignBriefSelectionTests.CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var card = PrepareCard(service, brief, assetPlan);

            using var start = new ManualResetEventSlim(false);
            var applyTask = Task.Run(() => CaptureRace("apply", () =>
            {
                start.Wait();
                return service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId);
            }));
            var cancelTask = Task.Run(() => CaptureRace("cancel", () =>
            {
                start.Wait();
                return service.CancelPendingSelection(Caller);
            }));

            start.Set();
            var attempts = await Task.WhenAll(applyTask, cancelTask);
            var winner = Assert.Single(attempts, static attempt => attempt.Value is not null);
            var rejected = Assert.Single(attempts, static attempt => attempt.ErrorCode is not null);

            if (winner.Operation == "apply")
            {
                Assert.IsType<DesignBriefValidationView>(winner.Value);
                Assert.Equal("cancel", rejected.Operation);
                Assert.Equal("design_brief_choice_cancel_forbidden", rejected.ErrorCode);
            }
            else
            {
                var cancellation = Assert.IsType<DesignBriefSelectionCancellationView>(winner.Value);
                Assert.Equal("cancelled", cancellation.Status);
                Assert.Equal("apply", rejected.Operation);
                Assert.Equal("design_brief_action_not_found", rejected.ErrorCode);
            }
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void CancelIsConversationBoundAndDoesNotAffectAnotherConversation()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, _) = DesignBriefSelectionTests.CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var card = PrepareCard(service, brief, assetPlan);
            var otherConversation = new CallerContext(
                "phase2-user",
                "other-conversation",
                "other-message");
            var otherUser = new CallerContext(
                "other-user",
                "phase2-conversation",
                "other-message");

            AssertCode(
                "design_brief_choice_not_pending",
                () => service.CancelPendingSelection(otherConversation));
            AssertCode(
                "design_brief_choice_not_pending",
                () => service.CancelPendingSelection(otherUser));

            var selected = service.ApplyCardSelection(
                Caller,
                card.ChoiceSessionId,
                card.Recommended.OptionId);
            Assert.Equal(card.Recommended.Binding.BriefId, selected.BriefId);
            AssertCode(
                "design_brief_choice_cancel_forbidden",
                () => service.CancelPendingSelection(Caller));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ChoiceExpiresAtTheExactBoundary()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, clock) = DesignBriefSelectionTests.CreateService(
                root,
                lifetimeMinutes: 1);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var beforeExpiryCaller = Caller;
            var atExpiryCaller = new CallerContext("expiry-user", "expiry-conversation", "expiry-message");
            var beforeExpiryCard = PrepareCard(service, brief, assetPlan, beforeExpiryCaller);
            var atExpiryCard = PrepareCard(service, brief, assetPlan, atExpiryCaller);

            Assert.Equal(beforeExpiryCard.ExpiresAt, atExpiryCard.ExpiresAt);
            clock.Advance(TimeSpan.FromMinutes(1) - TimeSpan.FromTicks(1));
            Assert.NotNull(service.ApplyCardSelection(
                beforeExpiryCaller,
                beforeExpiryCard.ChoiceSessionId,
                beforeExpiryCard.Recommended.OptionId));

            clock.Advance(TimeSpan.FromTicks(1));
            Assert.Equal(atExpiryCard.ExpiresAt, clock.GetUtcNow());
            AssertCode(
                "design_brief_action_expired",
                () => service.ApplyCardSelection(
                    atExpiryCaller,
                    atExpiryCard.ChoiceSessionId,
                    atExpiryCard.Recommended.OptionId));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void StartReservationBlocksMutationAndReleaseRestoresSelectedFlow()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, _) = DesignBriefSelectionTests.CreateService(
                root,
                requireDesignBrief: false);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var card = PrepareCard(service, brief, assetPlan);
            var selected = service.ApplyCardSelection(
                Caller,
                card.ChoiceSessionId,
                card.Recommended.OptionId);
            var binding = service.AuthorizeForStart(
                Caller,
                selected.BriefId,
                2,
                "none",
                null,
                null)!;

            Assert.True(service.ReserveStart(Caller, binding));
            AssertCode(
                "design_brief_start_in_progress",
                () => service.Validate(Caller, brief, assetPlan));
            AssertCode(
                "design_brief_start_in_progress",
                () => service.Prepare(
                    Caller,
                    brief,
                    assetPlan,
                    CreateAlternative(),
                    null,
                    replacePendingChoice: true));
            AssertCode(
                "design_brief_action_start_in_progress",
                () => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId));
            AssertCode(
                "design_brief_choice_cancel_forbidden",
                () => service.CancelPendingSelection(Caller));
            AssertCode(
                "design_brief_start_in_progress",
                () => service.ReserveStart(Caller, binding));

            service.ReleaseStartReservation(Caller, binding);

            Assert.Equal(
                selected.BriefId,
                service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId).BriefId);
            Assert.True(service.ReserveStart(Caller, binding));
            service.ReleaseStartReservation(Caller, binding);

            var optionalCaller = new CallerContext("optional-user", "optional-conversation", "optional-message");
            Assert.Null(service.AuthorizeForStart(optionalCaller, null, 2, "none", null, null));
            Assert.True(service.ReserveStart(optionalCaller, null));
            AssertCode(
                "design_brief_start_in_progress",
                () => service.Validate(optionalCaller, brief, assetPlan));
            AssertCode(
                "design_brief_start_in_progress",
                () => service.Prepare(optionalCaller, brief, assetPlan, CreateAlternative(), null));
            service.ReleaseStartReservation(optionalCaller, null);
            Assert.NotNull(service.Validate(optionalCaller, brief, assetPlan));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ConcurrentAuthorizeAndReserveAllowsOneStartAndRecoversAfterRelease()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, _) = DesignBriefSelectionTests.CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var card = PrepareCard(service, brief, assetPlan);
            var selected = service.ApplyCardSelection(
                Caller,
                card.ChoiceSessionId,
                card.Recommended.OptionId);

            using var start = new ManualResetEventSlim(false);
            var tasks = Enumerable.Range(0, 2).Select(_ => Task.Run(() =>
            {
                var binding = service.AuthorizeForStart(
                    Caller,
                    selected.BriefId,
                    2,
                    "none",
                    null,
                    null)!;
                start.Wait();
                try
                {
                    return new ReserveAttempt(binding, service.ReserveStart(Caller, binding), null);
                }
                catch (PptxValidationException exception)
                {
                    return new ReserveAttempt(binding, false, exception.Code);
                }
            })).ToArray();

            start.Set();
            var attempts = await Task.WhenAll(tasks);
            var winner = Assert.Single(attempts, static attempt => attempt.Reserved);
            var rejected = Assert.Single(attempts, static attempt => attempt.ErrorCode is not null);
            Assert.Equal("design_brief_start_in_progress", rejected.ErrorCode);

            service.ReleaseStartReservation(Caller, winner.Binding);
            var retry = service.AuthorizeForStart(
                Caller,
                selected.BriefId,
                2,
                "none",
                null,
                null)!;
            Assert.True(service.ReserveStart(Caller, retry));
            service.ReleaseStartReservation(Caller, retry);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void StartedSelectionRejectsStaleActionAndNewWorkflowButAllowsSameDraftRetry()
    {
        var root = CreateProfileRoot();
        try
        {
            var (service, catalog, _) = DesignBriefSelectionTests.CreateService(root);
            var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var card = PrepareCard(service, brief, assetPlan);
            var selected = service.ApplyCardSelection(
                Caller,
                card.ChoiceSessionId,
                card.Recommended.OptionId);
            var binding = service.AuthorizeForStart(
                Caller,
                selected.BriefId,
                2,
                "none",
                null,
                null)!;
            Assert.True(service.ReserveStart(Caller, binding));
            service.MarkStartSucceeded(Caller, binding.BriefId);

            AssertCode(
                "design_brief_action_already_started",
                () => service.ApplyCardSelection(
                    Caller,
                    card.ChoiceSessionId,
                    card.Recommended.OptionId));
            AssertCode(
                "design_brief_choice_cancel_forbidden",
                () => service.CancelPendingSelection(Caller));
            AssertCode(
                "design_brief_already_started",
                () => service.AuthorizeForStart(
                    Caller,
                    selected.BriefId,
                    2,
                    "none",
                    null,
                    null,
                    userRequestedNewWorkflow: true));

            var sameDraftRetry = service.AuthorizeForStart(
                Caller,
                selected.BriefId,
                2,
                "none",
                null,
                null)!;
            Assert.True(service.ReserveStart(Caller, sameDraftRetry));
            service.MarkStartSucceeded(Caller, sameDraftRetry.BriefId);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private static string CreateProfileRoot()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        DesignBriefSelectionTests.AddAlternativeDirection(root, materiallyDifferent: true);
        return root;
    }

    private static PreparedDesignBriefCard PrepareCard(
        DesignBriefService service,
        DesignBriefSpec brief,
        IReadOnlyList<AssetPlanItem> assetPlan,
        CallerContext? caller = null) =>
        service.Prepare(caller ?? Caller, brief, assetPlan, CreateAlternative(), null);

    private static DesignBriefStyleAlternative[] CreateAlternative() =>
    [
        new DesignBriefStyleAlternative(
            "report",
            "detailed",
            ["cover-report-airy", "kpi-report-detailed"]),
    ];

    private static void AssertCode(string expected, Action action)
    {
        var error = Assert.Throws<PptxValidationException>(action);
        Assert.Equal(expected, error.Code);
    }

    private static RaceAttempt CaptureRace(string operation, Func<object> action)
    {
        try
        {
            return new RaceAttempt(operation, action(), null);
        }
        catch (PptxValidationException exception)
        {
            return new RaceAttempt(operation, null, exception.Code);
        }
    }

    private sealed record ApplyAttempt(
        PreparedDesignBriefOption Option,
        DesignBriefValidationView? View,
        string? ErrorCode);

    private sealed record ReserveAttempt(
        ValidatedDesignBriefBinding Binding,
        bool Reserved,
        string? ErrorCode);

    private sealed record RaceAttempt(
        string Operation,
        object? Value,
        string? ErrorCode);
}
