using System.ComponentModel;
using System.Reflection;
using System.Text.Json;
using ModelContextProtocol.Server;
using PptxMcp.Domain;
using PptxMcp.Jobs;
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
            nameof(PowerPointTools.FinishBrandedVisualDeckAsync),
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
        var startMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.StartVisualDeck),
            BindingFlags.Public | BindingFlags.Static);
        var visualMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.FinishVisualDeckAsync),
            BindingFlags.Public | BindingFlags.Static);
        var brandedMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.FinishBrandedVisualDeckAsync),
            BindingFlags.Public | BindingFlags.Static);
        var strictMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.CreateDeckAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(startMethod);
        Assert.NotNull(visualMethod);
        Assert.NotNull(brandedMethod);
        Assert.NotNull(strictMethod);
        Assert.Equal(
            "default",
            startMethod.GetParameters().Single(parameter => parameter.Name == "templateSourceFileId").DefaultValue);
        Assert.Null(
            visualMethod.GetParameters().Single(parameter => parameter.Name == "useDefaultTemplate").DefaultValue);
        Assert.Null(brandedMethod.GetParameters().Single(parameter => parameter.Name == "sourceFileId").DefaultValue);
        Assert.Equal(
            "default",
            strictMethod.GetParameters().Single(parameter => parameter.Name == "sourceFileId").DefaultValue);
    }

    [Theory]
    [InlineData(nameof(PowerPointTools.CreateDeckAsync), "slides")]
    [InlineData(nameof(PowerPointTools.StartVisualDeck), "title")]
    [InlineData(nameof(PowerPointTools.StartVisualDeck), "expectedSlideCount")]
    [InlineData(nameof(PowerPointTools.AddVisualSlidesToDraft), "draftId")]
    [InlineData(nameof(PowerPointTools.AddVisualSlidesToDraft), "slides")]
    [InlineData(nameof(PowerPointTools.FinishVisualDeckAsync), "draftId")]
    [InlineData(nameof(PowerPointTools.FinishBrandedVisualDeckAsync), "draftId")]
    [InlineData(nameof(PowerPointTools.InsertVisualSlidesAsync), "slides")]
    [InlineData(nameof(PowerPointTools.RefineDeckAsync), "jobId")]
    [InlineData(nameof(PowerPointTools.RefineDeckAsync), "revisions")]
    [InlineData(nameof(PowerPointTools.RefineVisualDeckAsync), "jobId")]
    [InlineData(nameof(PowerPointTools.RefineVisualDeckAsync), "revisions")]
    public void BehaviorallyRequiredInputsAreRequiredInMcpSchema(string methodName, string argumentName)
    {
        var method = typeof(PowerPointTools).GetMethod(
            methodName,
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var tool = McpServerTool.Create(method);
        var required = tool.ProtocolTool.InputSchema
            .GetProperty("required")
            .EnumerateArray()
            .Select(element => element.GetString())
            .ToArray();

        Assert.Contains(argumentName, required);
    }

    [Fact]
    public void VisualDraftAppendPositionIsOptionalInMcpSchema()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.AddVisualSlidesToDraft),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var tool = McpServerTool.Create(method);
        var required = tool.ProtocolTool.InputSchema
            .GetProperty("required")
            .EnumerateArray()
            .Select(element => element.GetString())
            .ToArray();

        Assert.DoesNotContain("startSlideNumber", required);
        Assert.Null(method.GetParameters().Single(parameter => parameter.Name == "startSlideNumber").DefaultValue);
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
    public void VisualSlideInsertionDefaultsToLatestJobAndAppend()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.InsertVisualSlidesAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var description = method.GetCustomAttribute<DescriptionAttribute>()?.Description;
        var slides = method.GetParameters().Single(parameter => parameter.Name == "slides");
        var afterSlideNumber = method.GetParameters().Single(parameter => parameter.Name == "afterSlideNumber");
        var jobId = method.GetParameters().Single(parameter => parameter.Name == "jobId");

        Assert.Null(slides.DefaultValue);
        Assert.Null(afterSlideNumber.DefaultValue);
        Assert.Equal("latest", jobId.DefaultValue);
        Assert.Contains("追加分", description, StringComparison.Ordinal);
        Assert.Contains("資料全体を作り直さず", description, StringComparison.Ordinal);
        Assert.Contains("既存の資料タイトル、テーマ、design、全ページ、企業テンプレート", description, StringComparison.Ordinal);
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
    public void VisualDeckCreationUsesBoundedDraftBatchesInsteadOfOneLargeDeckArgument()
    {
        var start = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.StartVisualDeck),
            BindingFlags.Public | BindingFlags.Static);
        var add = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.AddVisualSlidesToDraft),
            BindingFlags.Public | BindingFlags.Static);
        var finish = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.FinishVisualDeckAsync),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(start);
        Assert.NotNull(add);
        Assert.NotNull(finish);
        Assert.DoesNotContain(start.GetParameters(), parameter => parameter.Name == "deck");
        Assert.DoesNotContain(finish.GetParameters(), parameter => parameter.Name == "deck");
        Assert.Contains("1〜4ページ", add.GetCustomAttribute<DescriptionAttribute>()?.Description, StringComparison.Ordinal);
        Assert.Equal(4, VisualDeckDraftService.MaximumBatchSlides);
    }

    [Fact]
    public void VisualDraftSchemaExposesCurrentStructuredAndEditableVocabulary()
    {
        var method = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.AddVisualSlidesToDraft),
            BindingFlags.Public | BindingFlags.Static);

        Assert.NotNull(method);
        var schema = McpServerTool.Create(method).ProtocolTool.InputSchema.GetRawText();
        var description = method.GetCustomAttribute<DescriptionAttribute>()?.Description;

        Assert.Contains("StructuredBrief", schema, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Scorecard", schema, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("MusicScore", schema, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("DataTable", schema, StringComparison.Ordinal);
        Assert.Contains("sections", schema, StringComparison.Ordinal);
        Assert.Contains("criteria", schema, StringComparison.Ordinal);
        Assert.Contains("musicScore", schema, StringComparison.Ordinal);
        Assert.Contains("dataTable", schema, StringComparison.Ordinal);
        Assert.Contains("recipeId", schema, StringComparison.Ordinal);
        Assert.Contains("density", schema, StringComparison.Ordinal);
        Assert.Contains("measures", schema, StringComparison.Ordinal);
        Assert.Contains("pitch", schema, StringComparison.Ordinal);
        Assert.Contains("fret", schema, StringComparison.Ordinal);
        Assert.Contains("finger", schema, StringComparison.Ordinal);
        Assert.Contains("structuredBrief", description, StringComparison.Ordinal);
        Assert.Contains("評価軸×選択肢", description, StringComparison.Ordinal);
        Assert.Contains("musicScore", description, StringComparison.Ordinal);
        Assert.Contains("五線譜", description, StringComparison.Ordinal);
        Assert.Contains("DataTable", description, StringComparison.Ordinal);
        Assert.Contains("dataTable", description, StringComparison.Ordinal);
        Assert.Contains("明示改行なし", description, StringComparison.Ordinal);
        Assert.Contains("recipeId", description, StringComparison.Ordinal);
        Assert.Contains("density", description, StringComparison.Ordinal);
        Assert.Contains("spotlight", description, StringComparison.Ordinal);
        Assert.Contains("Metrics正確に3件", description, StringComparison.Ordinal);
    }

    [Fact]
    public void OneCallVisualDeckCreationToolsAreNotExposed()
    {
        var toolNames = typeof(PowerPointTools)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(method => method.GetCustomAttribute<McpServerToolAttribute>()?.Name)
            .Where(name => name is not null)
            .ToArray();

        Assert.DoesNotContain("pptx_create_visual_deck", toolNames);
        Assert.DoesNotContain("pptx_create_branded_visual_deck", toolNames);
        Assert.Contains("pptx_start_visual_deck", toolNames);
        Assert.Contains("pptx_add_visual_slides_to_draft", toolNames);
        Assert.Contains("pptx_finish_visual_deck", toolNames);
        Assert.Contains("pptx_finish_branded_visual_deck", toolNames);
    }

    [Fact]
    public async Task InsertVisualSlidesReturnsActionableInputRequestForEmptyBedrockCall()
    {
        var result = await PowerPointTools.InsertVisualSlidesAsync(
            null!,
            null!,
            CancellationToken.None);

        var inputRequest = Assert.IsType<ToolInputRequest>(result);
        Assert.Equal("input_required", inputRequest.Status);
        Assert.Equal("pptx_insert_visual_slides", inputRequest.Tool);
        Assert.Equal(["slides"], inputRequest.RequiredArguments);
        Assert.Contains("only the new slides", inputRequest.Instruction, StringComparison.Ordinal);
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
        Assert.Contains("recipeId", toolDescription, StringComparison.Ordinal);
        Assert.Contains("実効density", toolDescription, StringComparison.Ordinal);

        var legacyMethod = typeof(PowerPointTools).GetMethod(
            nameof(PowerPointTools.RefineVisualDeckAsync),
            BindingFlags.Public | BindingFlags.Static);
        var legacyDescription = legacyMethod?.GetCustomAttribute<DescriptionAttribute>()?.Description;
        Assert.Contains("recipeId", legacyDescription, StringComparison.Ordinal);
        Assert.Contains("実効density", legacyDescription, StringComparison.Ordinal);
    }
}
