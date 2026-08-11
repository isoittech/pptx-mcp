using System.Reflection;
using System.Text.Json;
using ModelContextProtocol.Server;
using PptxMcp.Domain;
using PptxMcp.Tools;

namespace PptxMcp.Tests;

public sealed class DesignToolContractTests
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
        Assert.Contains("無引数ではcompactなprofile一覧だけ", tool.ProtocolTool.Description, StringComparison.Ordinal);
        Assert.Contains("profileIdと任意のpurpose", tool.ProtocolTool.Description, StringComparison.Ordinal);
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
