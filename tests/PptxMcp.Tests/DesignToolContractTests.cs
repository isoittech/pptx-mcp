using System.Reflection;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Server;
using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Security;
using PptxMcp.Storage;
using PptxMcp.Tools;
using Xunit.Abstractions;

namespace PptxMcp.Tests;

public sealed class DesignToolContractTests(ITestOutputHelper output)
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);

    [Fact]
    public void DesignCatalogUsesOneCompactFilterableTool()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.GetDesignCatalog),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var tool = McpServerTool.Create(method);
        var schema = tool.ProtocolTool.InputSchema;
        var required = schema.GetProperty("required")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .ToArray();

        Assert.Equal("pptx_get_design_catalog", tool.ProtocolTool.Name);
        Assert.DoesNotContain("profileId", required);
        Assert.DoesNotContain("purpose", required);
        Assert.DoesNotContain("density", required);
        Assert.DoesNotContain("styleDirectionId", required);
        Assert.Contains("profileId", schema.GetProperty("properties").EnumerateObject().Select(static item => item.Name));
        Assert.Contains("purpose", schema.GetProperty("properties").EnumerateObject().Select(static item => item.Name));
        Assert.Contains("density", schema.GetProperty("properties").EnumerateObject().Select(static item => item.Name));
        Assert.Contains("styleDirectionId", schema.GetProperty("properties").EnumerateObject().Select(static item => item.Name));
        Assert.Contains("正確に2回だけ", tool.ProtocolTool.Description, StringComparison.Ordinal);
        Assert.Contains("style_directions", tool.ProtocolTool.Description, StringComparison.Ordinal);
        Assert.Contains("複数用途", tool.ProtocolTool.Description, StringComparison.Ordinal);
        Assert.Contains("2回目の後は再呼出しません", tool.ProtocolTool.Description, StringComparison.Ordinal);
    }

    [Fact]
    public void BriefValidationRequiresCompleteBriefAndAssetPlan()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.ValidateDesignBrief),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var tool = McpServerTool.Create(method);
        var required = tool.ProtocolTool.InputSchema.GetProperty("required")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .ToArray();

        Assert.Equal("pptx_validate_design_brief", tool.ProtocolTool.Name);
        Assert.Contains("brief", required);
        Assert.Contains("assetPlan", required);
        Assert.Contains("acquisition=none, fallback=none, status=omitted", tool.ProtocolTool.Description, StringComparison.Ordinal);
        var schemaJson = tool.ProtocolTool.InputSchema.GetRawText();
        Assert.Contains("acquisition=none requires fallback=none", schemaJson, StringComparison.Ordinal);
        Assert.Contains("noAssetLayout", schemaJson, StringComparison.Ordinal);
        Assert.Contains("license_status=notRequired", schemaJson, StringComparison.Ordinal);
        Assert.DoesNotContain(method.GetParameters(), parameter =>
            parameter.Name is "url" or "path" or "sourceUrl" or "localPath");
    }

    [Fact]
    public void StartBriefIdRemainsSchemaOptionalForOssCompatibility()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.StartVisualDeck),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var tool = McpServerTool.Create(method);
        var required = tool.ProtocolTool.InputSchema.GetProperty("required")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .ToArray();

        Assert.DoesNotContain("briefId", required);
        Assert.Null(method.GetParameters().Single(parameter => parameter.Name == "briefId").DefaultValue);
    }

    [Fact]
    public void DesignBriefCardToolsExposeOnlyCompactBoundActionContract()
    {
        var prepareMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.PrepareDesignBrief),
            BindingFlags.Public | BindingFlags.Static);
        var applyMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.ApplyDesignBriefAction),
            BindingFlags.Public | BindingFlags.Static);
        var cancelMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.CancelDesignBriefSelection),
            BindingFlags.Public | BindingFlags.Static);
        Assert.NotNull(prepareMethod);
        Assert.NotNull(applyMethod);
        Assert.NotNull(cancelMethod);

        using var services = CreateToolServices();
        var createOptions = new McpServerToolCreateOptions { Services = services };
        var prepare = McpServerTool.Create(prepareMethod, (object)null!, createOptions).ProtocolTool;
        var apply = McpServerTool.Create(applyMethod, (object)null!, createOptions).ProtocolTool;
        var cancel = McpServerTool.Create(cancelMethod, (object)null!, createOptions).ProtocolTool;
        var prepareRequired = prepare.InputSchema.GetProperty("required")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .OfType<string>()
            .ToArray();
        var applyRequired = apply.InputSchema.GetProperty("required")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .OfType<string>()
            .ToArray();
        var applyProperties = apply.InputSchema.GetProperty("properties")
            .EnumerateObject()
            .Select(static item => item.Name)
            .ToArray();

        Assert.Equal("pptx_prepare_design_brief", prepare.Name);
        Assert.Equal("pptx_apply_design_brief_action", apply.Name);
        Assert.Equal("pptx_cancel_design_brief_selection", cancel.Name);
        Assert.Equal(["brief", "assetPlan"], prepareRequired);
        Assert.Equal(["choiceSessionId", "optionId"], applyRequired);
        Assert.Equal(["choiceSessionId", "optionId"], applyProperties);
        Assert.Empty(cancel.InputSchema.GetProperty("properties").EnumerateObject());
        Assert.DoesNotContain("briefId", applyProperties);
        Assert.DoesNotContain("nonce", applyProperties);
        Assert.DoesNotContain("action", applyProperties);
        Assert.DoesNotContain("styleDirectionId", applyProperties);
        Assert.Contains("2〜3件", prepare.Description, StringComparison.Ordinal);
        Assert.Contains("pptx.designBrief.select", apply.Description, StringComparison.Ordinal);
        Assert.DoesNotContain(
            prepareMethod.GetParameters().Concat(applyMethod.GetParameters()),
            parameter => parameter.Name is "url" or "path" or "sourceUrl" or "localPath");
    }

    [Fact]
    public void CardToolSchemaDeltaStaysWithinCompactRegressionBudget()
    {
        static int ContractBytes(string methodName)
        {
            var method = typeof(PowerPointTools).GetMethod(
                methodName,
                BindingFlags.Public | BindingFlags.Static);
            Assert.NotNull(method);
            using var services = CreateToolServices();
            var tool = McpServerTool.Create(
                method,
                (object)null!,
                new McpServerToolCreateOptions { Services = services }).ProtocolTool;
            return Encoding.UTF8.GetByteCount(tool.Description ?? string.Empty)
                + Encoding.UTF8.GetByteCount(tool.InputSchema.GetRawText());
        }

        var validateBytes = ContractBytes(nameof(PowerPointTools.ValidateDesignBrief));
        var prepareBytes = ContractBytes(nameof(PowerPointTools.PrepareDesignBrief));
        var applyBytes = ContractBytes(nameof(PowerPointTools.ApplyDesignBriefAction));
        var cancelBytes = ContractBytes(nameof(PowerPointTools.CancelDesignBriefSelection));
        output.WriteLine(
            $"validate={validateBytes}, prepare={prepareBytes}, apply={applyBytes}, cancel={cancelBytes}, phase2_total={prepareBytes + applyBytes + cancelBytes}");

        Assert.True(prepareBytes <= validateBytes + 6_000,
            $"prepare={prepareBytes}, validate={validateBytes}");
        Assert.True(applyBytes <= 2_500, $"apply={applyBytes}");
        Assert.True(cancelBytes <= 2_500, $"cancel={cancelBytes}");
        Assert.True(prepareBytes + applyBytes + cancelBytes <= 22_500,
            $"prepare={prepareBytes}, apply={applyBytes}, cancel={cancelBytes}, validate={validateBytes}");
    }

    private static ServiceProvider CreateToolServices() => new ServiceCollection()
        .AddSingleton<CallerContextAccessor>(_ => throw new InvalidOperationException("Schema-only service."))
        .AddSingleton<DesignBriefService>(_ => throw new InvalidOperationException("Schema-only service."))
        .AddSingleton<VisualObjectAssetRepository>(_ => throw new InvalidOperationException("Schema-only service."))
        .BuildServiceProvider();

    [Fact]
    public void VisualObjectToolIsBoundedSemanticAndCompact()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.PrepareVisualObjects),
            BindingFlags.Public | BindingFlags.Static);
        Assert.NotNull(method);
        using var services = CreateToolServices();
        var tool = McpServerTool.Create(
            method,
            (object)null!,
            new McpServerToolCreateOptions { Services = services }).ProtocolTool;
        var schema = tool.InputSchema.GetRawText();
        var bytes = Encoding.UTF8.GetByteCount(tool.Description ?? string.Empty)
            + Encoding.UTF8.GetByteCount(schema);

        Assert.Equal("pptx_prepare_visual_objects", tool.Name);
        Assert.Contains("objects", tool.InputSchema.GetProperty("required").EnumerateArray().Select(static item => item.GetString()));
        Assert.Contains("visualPurpose", schema, StringComparison.Ordinal);
        Assert.Contains("placementRole", schema, StringComparison.Ordinal);
        Assert.Contains("paletteRole", schema, StringComparison.Ordinal);
        Assert.DoesNotContain("svg", tool.InputSchema.GetProperty("properties").EnumerateObject().Select(static item => item.Name));
        Assert.DoesNotContain("url", tool.InputSchema.GetProperty("properties").EnumerateObject().Select(static item => item.Name));
        Assert.DoesNotContain("path", tool.InputSchema.GetProperty("properties").EnumerateObject().Select(static item => item.Name));
        Assert.True(bytes <= 9_000, $"visual object tool contract is {bytes} bytes");
    }

    [Fact]
    public void DesignEnumsSerializeAsStableCamelCaseTokens()
    {
        var plan = new AssetPlanItem(
            1,
            "cover",
            "cover-airy",
            AssetVisualPurpose.None,
            AssetPreferredMedium.None,
            AssetAcquisition.NativeDraw,
            AssetFallback.NoAssetLayout,
            AssetPlanStatus.FallbackSelected,
            AssetLicenseStatus.NotRequired);

        var json = JsonSerializer.Serialize(plan, SerializerOptions);

        Assert.Contains("\"acquisition\":\"nativeDraw\"", json, StringComparison.Ordinal);
        Assert.Contains("\"fallback\":\"noAssetLayout\"", json, StringComparison.Ordinal);
        Assert.Contains("\"license_status\":\"notRequired\"", json, StringComparison.Ordinal);
    }
}
