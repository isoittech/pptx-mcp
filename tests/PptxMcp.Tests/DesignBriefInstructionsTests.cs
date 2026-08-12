using PptxMcp.Configuration;
using PptxMcp.Tools;

namespace PptxMcp.Tests;

public sealed class DesignBriefInstructionsTests
{
    [Fact]
    public void RequiredGateOrdersCatalogValidationAndStart()
    {
        var instructions = PowerPointServerInstructions.Build(new PptxMcpOptions
        {
            RequireDesignBrief = true,
        });

        Assert.Contains("requires a validated Design Brief", instructions, StringComparison.Ordinal);
        Assert.Contains("call pptx_get_design_catalog first", instructions, StringComparison.Ordinal);
        Assert.Contains("Call pptx_validate_design_brief only after", instructions, StringComparison.Ordinal);
        Assert.Contains("call pptx_start_visual_deck with the returned briefId", instructions, StringComparison.Ordinal);
        Assert.Contains("preferred_medium=none, acquisition=none, fallback=none, status=omitted", instructions, StringComparison.Ordinal);
        Assert.Contains("Never pair acquisition=none with noAssetLayout", instructions, StringComparison.Ordinal);
        Assert.Contains("status=fallbackSelected with fallback=nativeDraw or noAssetLayout", instructions, StringComparison.Ordinal);
        Assert.Contains("match that recipe's semantic kind, density, and implemented variant", instructions, StringComparison.Ordinal);
        Assert.Contains("spotlight only for exactly three Metrics", instructions, StringComparison.Ordinal);
        Assert.Contains("one explicit line", instructions, StringComparison.Ordinal);
        Assert.Contains("pptx_prepare_design_brief", instructions, StringComparison.Ordinal);
        Assert.Contains("pptx.designBrief.select", instructions, StringComparison.Ordinal);
        Assert.Contains("pptx_cancel_design_brief_selection", instructions, StringComparison.Ordinal);
        Assert.Contains("argument-free", instructions, StringComparison.Ordinal);
        Assert.Contains("visual_draft_not_found", instructions, StringComparison.Ordinal);
        Assert.Contains("never call add or finish with that draftId again", instructions, StringComparison.Ordinal);
    }

    [Fact]
    public void OptionalGatePreservesExistingVisualDeckWorkflow()
    {
        var instructions = PowerPointServerInstructions.Build(new PptxMcpOptions());

        Assert.Contains("does not require a Design Brief", instructions, StringComparison.Ordinal);
        Assert.Contains("existing Visual Deck callers remain compatible", instructions, StringComparison.Ordinal);
        Assert.Contains("blocks an unbound start when RequireDesignBrief is false", instructions, StringComparison.Ordinal);
    }

    [Fact]
    public void OptionsRejectUnsafeCatalogPathAndBriefLifetime()
    {
        var relativePath = new PptxMcpOptions
        {
            BrandProfilesRoot = "relative-brand-profiles",
        };
        var invalidLifetime = new PptxMcpOptions
        {
            DesignBriefLifetimeMinutes = 121,
        };

        var pathError = Assert.Throws<InvalidOperationException>(() => relativePath.Validate(requireSecrets: false));
        var lifetimeError = Assert.Throws<InvalidOperationException>(() => invalidLifetime.Validate(requireSecrets: false));

        Assert.Contains("BrandProfilesRoot", pathError.Message, StringComparison.Ordinal);
        Assert.Contains("DesignBriefLifetimeMinutes", lifetimeError.Message, StringComparison.Ordinal);
    }
}
