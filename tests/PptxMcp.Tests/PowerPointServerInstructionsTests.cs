using PptxMcp.Configuration;
using PptxMcp.Tools;

namespace PptxMcp.Tests;

public sealed class PowerPointServerInstructionsTests
{
    [Fact]
    public void ConfiguredDefaultTemplateAvoidsRuntimeAnalysis()
    {
        var instructions = PowerPointServerInstructions.Build(new PptxMcpOptions
        {
            DefaultTemplateId = "organization-default",
        });

        Assert.Contains("without pptx_analyze", instructions, StringComparison.Ordinal);
        Assert.Contains("useDefaultTemplate=false", instructions, StringComparison.Ordinal);
        Assert.Contains("explicit alternate template overrides", instructions, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("call pptx_wait_for_job once", instructions, StringComparison.Ordinal);
        Assert.Contains("instead of repeatedly calling pptx_get_job", instructions, StringComparison.Ordinal);
    }

    [Fact]
    public void DeploymentNoticeIsInjectedVerbatimAndOnlyWhenConfigured()
    {
        const string notice = "PowerPoint資料には既定テンプレートを使用します。";

        var configured = PowerPointServerInstructions.Build(new PptxMcpOptions
        {
            FirstAssistantNotice = notice,
        });
        var unconfigured = PowerPointServerInstructions.Build(new PptxMcpOptions());

        Assert.Contains(notice, configured, StringComparison.Ordinal);
        Assert.Contains("one-time user notice", configured, StringComparison.Ordinal);
        Assert.Contains("Never emit the notice in a continuation produced after receiving any tool result", configured, StringComparison.Ordinal);
        Assert.DoesNotContain("one-time user notice", unconfigured, StringComparison.Ordinal);
    }
}
