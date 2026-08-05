using System.ComponentModel;
using System.Reflection;
using System.Text.Json;
using PptxMcp.Domain;
using PptxMcp.Tools;

namespace PptxMcp.Tests;

public sealed class ToolInputContractTests
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);

    [Fact]
    public void DeckInputUsesAnalysisCompatibleSnakeCaseKeys()
    {
        var slides = new[]
        {
            new DeckSlideSpec(
                "/ppt/slideLayouts/slideLayout1.xml",
                [new DeckField("Executive summary", ShapeId: 2, PlaceholderIndex: 7)]),
        };

        using var document = JsonDocument.Parse(JsonSerializer.Serialize(slides, SerializerOptions));
        var slide = document.RootElement[0];
        var field = slide.GetProperty("fields")[0];

        Assert.Equal("/ppt/slideLayouts/slideLayout1.xml", slide.GetProperty("layout_id").GetString());
        Assert.False(slide.TryGetProperty("layoutId", out _));
        Assert.Equal((uint)2, field.GetProperty("shape_id").GetUInt32());
        Assert.Equal((uint)7, field.GetProperty("placeholder_index").GetUInt32());
        Assert.False(field.TryGetProperty("shapeId", out _));
    }

    [Fact]
    public void DeckInputDeserializesAnalysisCompatibleSnakeCaseKeys()
    {
        const string json = """
            {
              "layout_id": "/ppt/slideLayouts/slideLayout27.xml",
              "fields": [
                { "text": "Market outlook", "shape_id": 2 }
              ]
            }
            """;

        var slide = JsonSerializer.Deserialize<DeckSlideSpec>(json, SerializerOptions);

        Assert.NotNull(slide);
        Assert.Equal("/ppt/slideLayouts/slideLayout27.xml", slide.LayoutId);
        Assert.Single(slide.Fields);
        Assert.Equal((uint)2, slide.Fields[0].ShapeId);
    }

    [Fact]
    public void DeckInputAcceptsLegacyCamelCaseDuringSchemaCacheMigration()
    {
        const string json = """
            {
              "layoutId": "/ppt/slideLayouts/slideLayout27.xml",
              "fields": [
                {
                  "text": "Market outlook",
                  "shapeId": 2,
                  "placeholderIndex": 7
                }
              ]
            }
            """;

        var slide = JsonSerializer.Deserialize<DeckSlideSpec>(json, SerializerOptions);

        Assert.NotNull(slide);
        Assert.Equal("/ppt/slideLayouts/slideLayout27.xml", slide.LayoutId);
        Assert.Single(slide.Fields);
        Assert.Equal((uint)2, slide.Fields[0].ShapeId);
        Assert.Equal((uint)7, slide.Fields[0].PlaceholderIndex);

        var canonicalJson = JsonSerializer.Serialize(slide, SerializerOptions);
        Assert.Contains("\"layout_id\"", canonicalJson, StringComparison.Ordinal);
        Assert.Contains("\"shape_id\"", canonicalJson, StringComparison.Ordinal);
        Assert.DoesNotContain("\"layoutId\"", canonicalJson, StringComparison.Ordinal);
        Assert.DoesNotContain("\"shapeId\"", canonicalJson, StringComparison.Ordinal);
    }

    [Fact]
    public void DeckInputSerializesStructuredParagraphSemantics()
    {
        var field = new DeckField(
            ShapeId: 5,
            Paragraphs:
            [
                new DeckParagraph("論点", DeckParagraphKind.Bullet),
                new DeckParagraph("実行", DeckParagraphKind.Numbered, Level: 1, StartAt: 2),
            ]);

        using var document = JsonDocument.Parse(JsonSerializer.Serialize(field, SerializerOptions));
        var paragraphs = document.RootElement.GetProperty("paragraphs");

        Assert.Equal("Bullet", paragraphs[0].GetProperty("kind").GetString());
        Assert.Equal("Numbered", paragraphs[1].GetProperty("kind").GetString());
        Assert.Equal(1, paragraphs[1].GetProperty("level").GetInt32());
        Assert.Equal(2, paragraphs[1].GetProperty("start_at").GetInt32());
    }

    [Fact]
    public void EditingInputsUseAnalysisCompatibleSnakeCaseKeys()
    {
        var replacement = new TextReplacement("Before", "After", 3, "Body", 12);
        var field = new TemplateField(4, "Updated", "Title", 8);

        var replacementJson = JsonSerializer.Serialize(replacement, SerializerOptions);
        var fieldJson = JsonSerializer.Serialize(field, SerializerOptions);

        Assert.Contains("\"slide_number\":3", replacementJson, StringComparison.Ordinal);
        Assert.Contains("\"shape_id\":12", replacementJson, StringComparison.Ordinal);
        Assert.Contains("\"slide_number\":4", fieldJson, StringComparison.Ordinal);
        Assert.Contains("\"shape_name\":\"Title\"", fieldJson, StringComparison.Ordinal);
    }

    [Fact]
    public void EditingInputsAcceptLegacyCamelCaseDuringSchemaCacheMigration()
    {
        const string replacementJson = """
            {
              "find": "Before",
              "replace": "After",
              "slideNumber": 3,
              "shapeName": "Body",
              "shapeId": 12
            }
            """;
        const string fieldJson = """
            {
              "slideNumber": 4,
              "text": "Updated",
              "shapeName": "Title",
              "shapeId": 8
            }
            """;

        var replacement = JsonSerializer.Deserialize<TextReplacement>(replacementJson, SerializerOptions);
        var field = JsonSerializer.Deserialize<TemplateField>(fieldJson, SerializerOptions);

        Assert.NotNull(replacement);
        Assert.Equal(3, replacement.SlideNumber);
        Assert.Equal("Body", replacement.ShapeName);
        Assert.Equal((uint)12, replacement.ShapeId);
        Assert.NotNull(field);
        Assert.Equal(4, field.SlideNumber);
        Assert.Equal("Title", field.ShapeName);
        Assert.Equal((uint)8, field.ShapeId);
    }

    [Fact]
    public void CreateDeckGuidanceRejectsSourceOnlyCallsAndNamesExactKeys()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.CreateDeckAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var toolDescription = method.GetCustomAttribute<DescriptionAttribute>()?.Description;
        var slidesDescription = method
            .GetParameters()
            .Single(parameter => parameter.Name == "slides")
            .GetCustomAttribute<DescriptionAttribute>()?
            .Description;

        Assert.Contains("slidesは必須", toolDescription, StringComparison.Ordinal);
        Assert.Contains("sourceFileIdだけで呼んではいけません", toolDescription, StringComparison.Ordinal);
        Assert.Contains("layout_id", slidesDescription, StringComparison.Ordinal);
        Assert.Contains("shape_id", slidesDescription, StringComparison.Ordinal);
        Assert.Contains("snake_case", slidesDescription, StringComparison.Ordinal);
    }

    [Fact]
    public void BrandedVisualDeckContractPreservesCorporateChromeAndSupportsAutoLayout()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.CreateBrandedVisualDeckAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var toolDescription = method.GetCustomAttribute<DescriptionAttribute>()?.Description;
        var layoutDescription = method
            .GetParameters()
            .Single(parameter => parameter.Name == "templateLayoutId")
            .GetCustomAttribute<DescriptionAttribute>()?
            .Description;

        Assert.Contains("マスター", toolDescription, StringComparison.Ordinal);
        Assert.Contains("ロゴ", toolDescription, StringComparison.Ordinal);
        Assert.Contains("フッター", toolDescription, StringComparison.Ordinal);
        Assert.Contains("auto", layoutDescription, StringComparison.Ordinal);

        var specification = new BrandedVisualDeckSpec(
            new VisualDeckSpec(
                "ブランド資料",
                [new VisualSlideSpec(VisualSlideKind.Title, "タイトル")]),
            "/ppt/slideLayouts/slideLayout14.xml");
        var json = JsonSerializer.Serialize(specification, SerializerOptions);

        Assert.Contains("\"template_layout_id\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("templateLayoutId", json, StringComparison.Ordinal);
    }

    [Fact]
    public void NewDeckToolsDefaultToDeploymentTemplate()
    {
        var visualMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.CreateVisualDeckAsync),
            BindingFlags.Public | BindingFlags.Static);
        var brandedMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.CreateBrandedVisualDeckAsync),
            BindingFlags.Public | BindingFlags.Static);
        var strictMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.CreateDeckAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(visualMethod);
        Assert.NotNull(brandedMethod);
        Assert.NotNull(strictMethod);
        Assert.Equal(
            true,
            visualMethod.GetParameters().Single(parameter => parameter.Name == "useDefaultTemplate").DefaultValue);
        Assert.Equal(
            "default",
            brandedMethod.GetParameters().Single(parameter => parameter.Name == "sourceFileId").DefaultValue);
        Assert.Equal(
            "default",
            strictMethod.GetParameters().Single(parameter => parameter.Name == "sourceFileId").DefaultValue);
    }

    [Fact]
    public void SingleSlideRefinementGuidanceAvoidsRedundantPolling()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.RefineVisualSlideAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var description = method.GetCustomAttribute<DescriptionAttribute>()?.Description;

        Assert.Contains("Succeededなら状態確認ツールを呼ばず", description, StringComparison.Ordinal);
        Assert.Contains("30秒以内に完了せずqueuedを返した場合だけ", description, StringComparison.Ordinal);
        Assert.Contains("pptx_wait_for_job", description, StringComparison.Ordinal);
    }

    [Fact]
    public void WaitForJobContractUsesBoundedServerSideWait()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.WaitForJobAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var description = method.GetCustomAttribute<DescriptionAttribute>()?.Description;
        var jobId = method.GetParameters().Single(parameter => parameter.Name == "jobId");
        var waitSeconds = method.GetParameters().Single(parameter => parameter.Name == "waitSeconds");

        Assert.Contains("短間隔ポーリング", description, StringComparison.Ordinal);
        Assert.Equal("latest", jobId.DefaultValue);
        Assert.Equal(45, waitSeconds.DefaultValue);
    }

    [Fact]
    public async Task WaitForJobRejectsOutOfRangeWaitBeforeAccessingServices()
    {
        var result = await PowerPointTools.WaitForJobAsync(
            null!,
            null!,
            CancellationToken.None,
            waitSeconds: 51);

        var error = Assert.IsType<ToolValidationError>(result);
        Assert.Equal("wait_seconds_invalid", error.Code);
        Assert.Equal("pptx_wait_for_job", error.Tool);
    }

    [Fact]
    public async Task CreateDeckReturnsActionableInputRequestForEmptyBedrockCall()
    {
        var result = await PowerPointTools.CreateDeckAsync(
            null!,
            null!,
            CancellationToken.None,
            slides: null);

        var inputRequest = Assert.IsType<ToolInputRequest>(result);
        Assert.Equal("input_required", inputRequest.Status);
        Assert.Equal("pptx_create_deck", inputRequest.Tool);
        Assert.Equal(["slides"], inputRequest.RequiredArguments);
    }

    [Fact]
    public async Task RefineDeckReturnsActionableInputRequestForEmptyBedrockCall()
    {
        var result = await PowerPointTools.RefineDeckAsync(
            null!,
            null!,
            CancellationToken.None);

        var inputRequest = Assert.IsType<ToolInputRequest>(result);
        Assert.Equal("input_required", inputRequest.Status);
        Assert.Equal("pptx_refine_deck", inputRequest.Tool);
        Assert.Equal(["jobId", "revisions"], inputRequest.RequiredArguments);
    }

    [Fact]
    public async Task RefineVisualDeckReturnsActionableInputRequestForEmptyBedrockCall()
    {
        var result = await PowerPointTools.RefineVisualDeckAsync(
            null!,
            null!,
            CancellationToken.None);

        var inputRequest = Assert.IsType<ToolInputRequest>(result);
        Assert.Equal("input_required", inputRequest.Status);
        Assert.Equal("pptx_refine_visual_deck", inputRequest.Tool);
        Assert.Equal(["jobId", "revisions"], inputRequest.RequiredArguments);
    }

    [Fact]
    public async Task CreateVisualDeckReturnsActionableInputRequestForEmptyBedrockCall()
    {
        var result = await PowerPointTools.CreateVisualDeckAsync(
            null!,
            null!,
            CancellationToken.None);

        var inputRequest = Assert.IsType<ToolInputRequest>(result);
        Assert.Equal("input_required", inputRequest.Status);
        Assert.Equal("pptx_create_visual_deck", inputRequest.Tool);
        Assert.Equal(["deck"], inputRequest.RequiredArguments);
    }

    [Fact]
    public async Task CreateBrandedVisualDeckReturnsActionableInputRequestForEmptyBedrockCall()
    {
        var result = await PowerPointTools.CreateBrandedVisualDeckAsync(
            null!,
            null!,
            CancellationToken.None);

        var inputRequest = Assert.IsType<ToolInputRequest>(result);
        Assert.Equal("input_required", inputRequest.Status);
        Assert.Equal("pptx_create_branded_visual_deck", inputRequest.Tool);
        Assert.Equal(["deck"], inputRequest.RequiredArguments);
    }

    [Fact]
    public void VisualValidationErrorsExposeTheExactFieldFailureToTheModel()
    {
        var error = new ToolValidationError(
            "invalid_input",
            "pptx_refine_visual_deck",
            "visual_content_missing",
            "slides[9].body is required for a statement slide.",
            "Correct the field and retry.");

        var json = JsonSerializer.Serialize(error, SerializerOptions);

        Assert.Contains("\"status\":\"invalid_input\"", json, StringComparison.Ordinal);
        Assert.Contains("\"code\":\"visual_content_missing\"", json, StringComparison.Ordinal);
        Assert.Contains("slides[9].body", json, StringComparison.Ordinal);
    }

    [Fact]
    public void SingleSlideRefinementIsRequiredAndDefaultsToLatestJob()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.RefineVisualSlideAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var revision = method.GetParameters().Single(parameter => parameter.Name == "revision");
        var jobId = method.GetParameters().Single(parameter => parameter.Name == "jobId");
        var toolDescription = method.GetCustomAttribute<DescriptionAttribute>()?.Description;

        Assert.False(revision.HasDefaultValue);
        Assert.Equal("latest", jobId.DefaultValue);
        Assert.Contains("1枚", toolDescription, StringComparison.Ordinal);
        Assert.Contains("累積", toolDescription, StringComparison.Ordinal);
    }
}
