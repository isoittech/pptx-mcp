using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Security;
using PptxMcp.Storage;

namespace PptxMcp.Tools;

[McpServerToolType]
public sealed class PowerPointTools
{
    [McpServerTool(Name = "pptx_get_capabilities", ReadOnly = true, Idempotent = true),
     Description("PowerPoint MCPの対応範囲、制約、推奨ワークフローを返します。")]
    public static object GetCapabilities() => new
    {
        workflow = new[]
        {
            "LibreChatへPPTXをアップロードする",
            "pptx_analyzeでスライドと編集候補を取得する（file_idが会話に提示されない場合はsourceFileIdを省略して最新アップロードを使う）",
            "対象が曖昧ならスライド番号とshape_idをユーザーに選択してもらう",
            "企業テンプレートから新規作成する場合は、解析結果のlayout_idとshape_idをそのまま使い、全ページ分のslidesを1回のpptx_create_deckへ渡す",
            "企業テンプレートのマスター、ロゴ、フッターを保ちながら華やかな資料を作る場合はpptx_create_branded_visual_deckを使う",
            "既存スライドを更新する場合はpptx_replace_textまたはpptx_populate_templateを使う",
            "白紙から作る場合はpptx_create_visual_deckで意味ベースのレイアウトとデザイン方針を指定する",
            "pptx_get_jobで完了を確認する",
            "pptx_get_preview_imagesで全ページを実際に見て、企業テンプレート資料はpptx_refine_deck、白紙・ブランドVisual Deckはpptx_refine_visual_slideへ問題ページを1枚ずつ渡して最大2巡まで再生成する",
            "視覚確認後にのみPPTXのダウンロードリンクを提示する",
        },
        limits = new
        {
            maximum_file_bytes = 30L * 1024 * 1024,
            maximum_slides = 50,
            maximum_concurrent_jobs = 3,
            maximum_job_minutes = 10,
        },
        supported_now = new[]
        {
            "PPTX構造・テキストシェイプ・テーマ色・日本語フォント・SmartArt/グラフ/埋め込みExcel有無の解析",
            "通常シェイプとSmartArt内部テキストの置換",
            "名前付きシェイプを持つ企業テンプレートへのテキスト流し込み",
            "定義済みlayout_idから1〜50枚の新規デッキを構成",
            "白紙からKPI・カード・比較・工程・タイムライン・マトリクス・ファネル・ロードマップ・ダッシュボード・編集可能グラフ等の視覚的なデッキを構成",
            "企業テンプレートの空白レイアウトと自動抽出したテーマをVisual Deckへ合成し、ブランド要素と編集可能な視覚表現を両立",
            "LibreOfficeによる全スライドPNGプレビュー",
            "プレビュー画像をClaudeへ返す自動視覚リフレクション",
            "署名付きURLによる成果物ダウンロード",
        },
        planned_after_template_spike = new[]
        {
            "SmartArtノードの追加・削除（種類変更は対象外）",
            "編集可能グラフのデータ・系列・凡例・軸・色の更新（種類変更は対象外）",
            "埋め込みExcelの値・式・範囲・行列更新",
            "既存デッキに対するスライド追加・削除・並べ替え",
        },
    };

    [McpServerTool(Name = "pptx_analyze", ReadOnly = true, Idempotent = true),
     Description("LibreChatにアップロード済みのPPTXを安全に検査し、編集候補、企業テーマ色、日本語フォントを非同期で解析します。file_idが不明ならsourceFileIdを省略し、そのユーザーの最新PPTXを使ってください。解析結果のthemeは白紙VisualDeckSpecのthemeへ移植できます。")]
    public static Task<JobReceipt> AnalyzeAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("LibreChatのアップロード済みPPTXのfile_id。会話にfile_idが提示されていなければ省略して最新アップロードを使います。ファイル名やパスは指定しません。")]
        string sourceFileId = "latest") =>
        jobs.SubmitAnalyzeAsync(callerContext.GetRequired(), sourceFileId, cancellationToken);

    [McpServerTool(Name = "pptx_render_preview", ReadOnly = true, Idempotent = true),
     Description("アップロード済みPPTXの全スライドをPNGへ変換する非同期ジョブを開始します。")]
    public static Task<JobReceipt> RenderPreviewAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("LibreChatのアップロード済みPPTXのfile_id。省略時はそのユーザーの最新PPTX。")]
        string sourceFileId = "latest") =>
        jobs.SubmitRenderAsync(callerContext.GetRequired(), sourceFileId, cancellationToken);

    [McpServerTool(Name = "pptx_replace_text", Destructive = true),
     Description("PPTXの文字列を置換し、更新版PPTXと全ページプレビューを作る非同期ジョブを開始します。曖昧な対象は先にpptx_analyzeで特定してください。")]
    public static Task<JobReceipt> ReplaceTextAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("検索文字、置換文字、任意の1始まりスライド番号、任意のshapeNameとshapeIdからなる置換指示。名前が重複する場合はshapeIdも指定します。")]
        IReadOnlyList<TextReplacement> replacements,
        CancellationToken cancellationToken,
        [Description("LibreChatのアップロード済みPPTXのfile_id。省略時はそのユーザーの最新PPTX。")]
        string sourceFileId = "latest") =>
        jobs.SubmitReplaceTextAsync(callerContext.GetRequired(), sourceFileId, replacements, cancellationToken);

    [McpServerTool(Name = "pptx_populate_template", Destructive = true),
     Description("企業テンプレートの既存スライドにある名前付きテキストシェイプへ内容を流し込み、PPTXと全ページプレビューを作ります。通常文はtext、箇条書き・番号付き手順はparagraphsを使い、記号や番号を本文へ手入力しないでください。")]
    public static Task<JobReceipt> PopulateTemplateAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("1始まりのスライド番号、解析で得たshapeNameまたはshapeId、textまたはparagraphsのどちらか一方からなるフィールド一覧。paragraphsは各項目にtext、kind=Plain/Bullet/Numbered、level=0〜4、番号開始時だけstart_atを指定します。名前が重複する場合はshapeIdを指定します。")]
        IReadOnlyList<TemplateField> fields,
        CancellationToken cancellationToken,
        [Description("LibreChatへアップロード済みのテンプレートPPTXのfile_id。省略時はそのユーザーの最新PPTX。")]
        string sourceFileId = "latest") =>
        jobs.SubmitPopulateTemplateAsync(callerContext.GetRequired(), sourceFileId, fields, cancellationToken);

    [McpServerTool(Name = "pptx_create_deck", Destructive = true),
     Description("企業テンプレートのマスターと定義済みレイアウトを使い、1〜50枚の新規PPTXと全ページプレビューを作ります。先にpptx_analyzeとpptx_get_jobでlayout_idとplaceholderのshape_idを取得してください。通常文はtext、箇条書き・番号付き手順はparagraphsを使い、記号や番号を本文へ手入力しません。slidesは必須で、完成版の全ページを1回の呼び出しに含めます。sourceFileIdだけで呼んではいけません。slidesが欠けた呼び出しにはinput_requiredを返すので、同じツールを全slides付きで直ちに再実行してください。各layout_id、shape_id、placeholder_indexは解析結果から一字も変更せずコピーし、推測したパスやIDを作らないでください。")]
    public static async Task<object> CreateDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("動作上必須。完成版の全1〜50ページからなる配列です。各要素はlayout_idとfields、各fieldはtextまたはparagraphsのどちらか一方と解析結果のshape_id（任意でshape_nameまたはplaceholder_index）を使います。paragraphsの各項目はtext、kind=Plain/Bullet/Numbered、level=0〜4、任意のstart_atです。例: [{\"layout_id\":\"/ppt/slideLayouts/slideLayout1.xml\",\"fields\":[{\"paragraphs\":[{\"text\":\"現状把握\",\"kind\":\"Numbered\",\"level\":0}],\"shape_id\":2}]}]。キーはsnake_caseを厳守し、全ページを組み立て終えてから1回だけ呼びます。")]
        IReadOnlyList<DeckSlideSpec>? slides = null,
        [Description("LibreChatへアップロード済みの企業テンプレートPPTXのfile_id。省略時はそのユーザーの最新PPTX。")]
        string sourceFileId = "latest")
    {
        if (slides is null || slides.Count == 0)
        {
            return new ToolInputRequest(
                "input_required",
                "pptx_create_deck",
                ["slides"],
                "Call pptx_create_deck again now with the complete 1-50 slide array. Do not call it with only sourceFileId.");
        }

        return await jobs.SubmitCreateDeckAsync(
            callerContext.GetRequired(),
            sourceFileId,
            slides,
            cancellationToken).ConfigureAwait(false);
    }

    [McpServerTool(Name = "pptx_refine_deck", Destructive = true),
     Description("成功したpptx_create_deckジョブの仕様を再利用し、視覚確認で問題があったページだけを差し替えて再生成します。全10ページ等の完全仕様を再送せず、変更ページだけをrevisionsへ指定してください。元のlayout_idは自動的に保持されます。jobIdまたはrevisionsが欠けた呼び出しにはinput_requiredを返すので、両方を付けて直ちに再実行してください。")]
    public static async Task<object> RefineDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("動作上必須。直前の成功したpptx_create_deckが返したjob_idをjobIdとして指定します。")]
        string jobId = "",
        [Description("動作上必須。変更するページだけの配列です。各要素は1始まりのslide_numberと、そのページに残す全フィールドのfieldsをsnake_caseで指定します。fieldsの各要素はtextと元ページで使ったshape_idを含めます。")]
        IReadOnlyList<DeckSlideRevision>? revisions = null)
    {
        var missing = new List<string>(2);
        var requestedRevisions = revisions ?? [];
        if (string.IsNullOrWhiteSpace(jobId))
        {
            missing.Add("jobId");
        }

        if (requestedRevisions.Count == 0)
        {
            missing.Add("revisions");
        }

        if (missing.Count > 0)
        {
            return new ToolInputRequest(
                "input_required",
                "pptx_refine_deck",
                missing,
                "Call pptx_refine_deck again now with the successful deck jobId and only the changed slides in revisions.");
        }

        return await jobs.SubmitRefineDeckAsync(
            callerContext.GetRequired(),
            jobId,
            requestedRevisions,
            cancellationToken).ConfigureAwait(false);
    }

    [McpServerTool(Name = "pptx_create_visual_deck", Destructive = true),
     Description("アップロード済みテンプレートを使わず、検証済みの宣言型レイアウトから華やかで編集可能な16:9 PPTXを作ります。企業テンプレートが添付され、ブランド要素も維持する依頼ではpptx_create_branded_visual_deckを優先してください。title/agenda/section/statement/cards/metrics/comparison/process/timeline/matrix/funnel/roadmap/chart/dashboard/quote/closingを内容に応じて使い分け、6枚以上では4種類以上の構図を使ってください。単純な箇条書きは補助的に限定し、各ページに図形・データ・工程・比較等の視覚的な主役を置きます。生成後はpptx_get_jobを待ち、必ずpptx_get_preview_imagesで全スライドを確認し、問題ページをpptx_refine_visual_slideで1枚ずつ最大2巡まで修正してください。")]
    public static async Task<object> CreateVisualDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("動作上必須。資料タイトル、任意のthemeとdesign、1〜50枚のスライドからなる宣言型仕様。Metricsはmetricsを2〜6件、Dashboardはmetricsを2〜4件とchart、Cardsはcardsを3〜6件、Comparisonはpanelsを2〜3件、Process/Timeline/Funnel/Roadmapはstepsを3〜6件、Matrixはquadrantsを正確に4件指定します。design.styleはexecutive/editorial/bold/technical/playful、densityはairy/balanced/detailed、motifはgeometric/orbit/nodes/ribbon/noneです。各スライドは1枚1メッセージに絞り、variant=auto/grid/spotlight/split/cascade/editorialで構図を選べます。metric/cardのtoneはaccent/positive/success/warning/critical/danger/negative/neutral/info等の意味語または#RRGGBBを使えます。iconはinsight/target/growth/people/shield/clock/cloud/settings/data/warning/check/idea/search/compliance/decision/lock/network/document/communication/recovery/backup/legal/monitor/automationです。theme.presetはmidnight/aurora/sunset/forest/minimal/ocean/berry/clay/cyberです。")]
        VisualDeckSpec? deck = null)
    {
        if (deck is null)
        {
            return new ToolInputRequest(
                "input_required",
                "pptx_create_visual_deck",
                ["deck"],
                "Call pptx_create_visual_deck again with a complete VisualDeckSpec in deck. Do not call with empty arguments.");
        }

        try
        {
            return await jobs.SubmitVisualDeckAsync(
                callerContext.GetRequired(),
                deck,
                cancellationToken).ConfigureAwait(false);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_create_visual_deck", exception);
        }
    }

    [McpServerTool(Name = "pptx_create_branded_visual_deck", Destructive = true),
     Description("アップロード済み企業テンプレートのマスター、ロゴ、フッター、ページ設定を維持し、テンプレートから自動抽出した色と日本語フォントを適用したVisual Deckを合成します。厳密な既存プレースホルダー流し込みではなく、企業ブランドを保ちながらカード、KPI、工程、タイムライン、マトリクス、ファネル、ロードマップ、ダッシュボード、編集可能グラフ等の華やかな資料を作る場合に使います。deckにはpptx_create_visual_deckと同じ完全なVisualDeckSpecを渡します。templateLayoutIdは通常autoを使い、テンプレート内のプレースホルダー0個の白紙（フッター有を優先）レイアウトを自動選択します。生成後は全ページを視覚確認し、問題ページをpptx_refine_visual_slideで1枚ずつ最大2巡まで修正してください。")]
    public static async Task<object> CreateBrandedVisualDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("動作上必須。資料タイトル、design、1〜50枚のスライドからなる完全なVisualDeckSpec。Metricsはmetricsを2〜6件、Dashboardはmetricsを2〜4件とchart、Cardsはcardsを3〜6件、Comparisonはpanelsを2〜3件、Process/Timeline/Funnel/Roadmapはstepsを3〜6件、Matrixはquadrantsを正確に4件指定します。themeの色とフォントは企業テンプレートの値で自動上書きされるため、モデルが解析値を転記する必要はありません。")]
        VisualDeckSpec? deck = null,
        [Description("通常はauto。明示する場合はpptx_analyzeが返したプレースホルダー0個の白紙layout_idを一字も変更せず指定します。")]
        string templateLayoutId = "auto",
        [Description("LibreChatへアップロード済みの企業テンプレートPPTXのfile_id。省略時はそのユーザーの最新PPTX。")]
        string sourceFileId = "latest")
    {
        if (deck is null)
        {
            return new ToolInputRequest(
                "input_required",
                "pptx_create_branded_visual_deck",
                ["deck"],
                "Call pptx_create_branded_visual_deck again with a complete VisualDeckSpec in deck. Keep sourceFileId and templateLayoutId if already known; do not call with empty arguments.");
        }

        try
        {
            return await jobs.SubmitBrandedVisualDeckAsync(
                callerContext.GetRequired(),
                sourceFileId,
                deck,
                templateLayoutId,
                cancellationToken).ConfigureAwait(false);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_create_branded_visual_deck", exception);
        }
    }

    [McpServerTool(Name = "pptx_refine_visual_deck", Destructive = true),
     Description("成功したpptx_create_visual_deckまたはpptx_create_branded_visual_deckジョブの仕様を再利用し、視覚確認で問題があったページだけを差し替えて再生成します。ブランドVisual Deckでは元の企業テンプレートと選択レイアウトも自動的に維持します。資料全体を再送せず、変更ページの完全なVisualSlideSpecだけをrevisionsへ指定してください。jobIdまたはrevisionsが欠けた場合はinput_requiredを返します。")]
    public static async Task<object> RefineVisualDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("動作上必須。直前の成功したpptx_create_visual_deckまたはpptx_create_branded_visual_deckが返したjob_id。")]
        string jobId = "",
        [Description("動作上必須。変更ページだけの配列。各要素は1始まりのslide_numberと、差し替え後の完全なslideを含めます。")]
        IReadOnlyList<VisualSlideRevision>? revisions = null)
    {
        var missing = new List<string>(2);
        var requestedRevisions = revisions ?? [];
        if (string.IsNullOrWhiteSpace(jobId))
        {
            missing.Add("jobId");
        }

        if (requestedRevisions.Count == 0)
        {
            missing.Add("revisions");
        }

        if (missing.Count > 0)
        {
            return new ToolInputRequest(
                "input_required",
                "pptx_refine_visual_deck",
                missing,
                "Call pptx_refine_visual_deck again with the successful visual deck jobId and only the changed slides in revisions.");
        }

        try
        {
            return await jobs.SubmitRefineVisualDeckAsync(
                callerContext.GetRequired(),
                jobId,
                requestedRevisions,
                cancellationToken).ConfigureAwait(false);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_refine_visual_deck", exception);
        }
    }

    [McpServerTool(Name = "pptx_refine_visual_slide", Destructive = true),
     Description("成功したVisual DeckまたはブランドVisual Deckの問題ページを1枚だけ差し替えます。Bedrock/Claudeでの自動視覚リフレクションでは、大きなrevisions配列を作るpptx_refine_visual_deckよりこのツールを優先してください。revisionにslide_numberと差し替え後の完全なslideを必ず含めます。jobIdは通常latestを使い、会話内の最新成功Visual Deckを自動選択します。通常はサーバー内で完了まで待って最終statusと成果物を返すため、Succeededならpptx_get_jobを呼ばず次のページへ進んでください。30秒以内に完了しない場合だけjob_idをpptx_get_jobで確認します。複数ページを1ページずつ直すと修正が累積します。全ページを最大2巡までに収束させてください。")]
    public static async Task<object> RefineVisualSlideAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("動作上必須。1始まりのslide_numberと、そのページを完全に置き換えるVisualSlideSpecであるslide。1回の呼び出しには必ず1ページだけを含めます。")]
        VisualSlideRevision revision,
        CancellationToken cancellationToken,
        [Description("通常はlatest。明示する場合は直前に成功したVisual Deckジョブのjob_id。latestは同じ会話の最新成功ジョブを安全に選びます。")]
        string jobId = "latest")
    {
        try
        {
            var caller = callerContext.GetRequired();
            var resolvedJobId = string.IsNullOrWhiteSpace(jobId)
                || string.Equals(jobId, "latest", StringComparison.OrdinalIgnoreCase)
                ? await jobs.GetLatestSuccessfulVisualJobIdAsync(caller, cancellationToken).ConfigureAwait(false)
                : jobId;
            var receipt = await jobs.SubmitRefineVisualDeckAsync(
                caller,
                resolvedJobId,
                [revision],
                cancellationToken).ConfigureAwait(false);
            var completed = await jobs.WaitForTerminalAsync(
                caller,
                receipt.JobId,
                TimeSpan.FromSeconds(30),
                cancellationToken).ConfigureAwait(false);
            return completed is not null ? completed : receipt;
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_refine_visual_slide", exception);
        }
    }

    [McpServerTool(Name = "pptx_get_job", ReadOnly = true, Idempotent = true),
     Description("PowerPointジョブの状態、解析結果、短時間有効なプレビューURLとダウンロードURLを取得します。jobId=latestでは同じ利用者・会話の直近ジョブ（待機中・実行中を含む）を安全に選びます。完了までpoll_after_secondsを目安に再実行してください。")]
    public static async Task<object> GetJobAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("PowerPoint MCPが返したjob_id、または同じ会話の直近ジョブを選ぶlatest。")]
        string jobId,
        CancellationToken cancellationToken)
    {
        try
        {
            return await jobs.GetAsync(callerContext.GetRequired(), jobId, cancellationToken).ConfigureAwait(false);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_get_job", exception);
        }
    }

    [McpServerTool(Name = "pptx_get_preview_images", ReadOnly = true, Idempotent = true),
     Description("成功済みジョブのスライド画像をClaude自身の視覚確認用に返します。全スライドを1〜4枚ずつ取得し、文字切れ・重なり・可読性に加え、単調な構図、文字中心、弱い視覚階層、余白、整列、コントラスト、密度、バランス、全体一貫性を確認してください。厳密なプレースホルダー資料はpptx_refine_deck、白紙またはブランドVisual Deckはpptx_refine_visual_slideへ問題ページを1枚ずつ渡します。最大2巡で収束させ、このツールを呼ばずに視覚確認済みと述べてはいけません。")]
    public static async Task<CallToolResult> GetPreviewImagesAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("PowerPoint MCPが返した成功済みjob_id。")]
        string jobId,
        [Description("今回確認する1始まりのスライド番号。重複なしで1〜4件。全ページを複数回に分けて指定します。")]
        IReadOnlyList<int> slideNumbers,
        CancellationToken cancellationToken)
    {
        var images = await jobs.GetPreviewImagesAsync(
            callerContext.GetRequired(),
            jobId,
            slideNumbers,
            cancellationToken).ConfigureAwait(false);
        var content = new List<ContentBlock>(images.Count * 2 + 1);
        foreach (var image in images)
        {
            content.Add(new TextContentBlock { Text = $"Slide {image.SlideNumber} visual review image:" });
            content.Add(ImageContentBlock.FromBytes(image.Bytes, image.MediaType));
        }

        content.Add(new TextContentBlock
        {
            Text = "Evaluate the returned slides for clipping, overflow, overlap, legibility, spacing, alignment, contrast, hierarchy, density, balance, visual variety, and consistency. For a placeholder template deck use pptx_refine_deck; for a visual or branded visual deck prefer pptx_refine_visual_slide with one complete replacement slide at a time. Never resend the complete deck.",
        });
        return new CallToolResult
        {
            Content = content,
            StructuredContent = JsonSerializer.SerializeToElement(new
            {
                job_id = jobId,
                reviewed_slide_numbers = images.Select(image => image.SlideNumber).ToArray(),
            }),
        };
    }

    [McpServerTool(Name = "pptx_cancel_job", Destructive = true),
     Description("待機中または実行中のPowerPointジョブをキャンセルします。")]
    public static Task<bool> CancelJobAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("PowerPoint MCPが返したjob_id。")]
        string jobId,
        CancellationToken cancellationToken) =>
        jobs.CancelAsync(callerContext.GetRequired(), jobId, cancellationToken);

    private static ToolValidationError CreateValidationError(string tool, PptxValidationException exception) =>
        new(
            "invalid_input",
            tool,
            exception.Code,
            exception.Message,
            $"Correct the field named in message and call {tool} again. Do not repeat the same invalid input.");
}
