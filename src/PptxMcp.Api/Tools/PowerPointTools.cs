using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Security;

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
            "既存スライドを更新する場合はpptx_replace_textまたはpptx_populate_templateを使う",
            "白紙から作る場合はpptx_create_visual_deckで意味ベースのレイアウトを指定する",
            "pptx_get_jobで完了を確認する",
            "pptx_get_preview_imagesで全ページを実際に見て、企業テンプレート資料の違和感はpptx_refine_deckへ変更ページだけを渡して最大2回まで再生成する",
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
            "PPTX構造・テキストシェイプ・SmartArt/グラフ/埋め込みExcel有無の解析",
            "通常シェイプとSmartArt内部テキストの置換",
            "名前付きシェイプを持つ企業テンプレートへのテキスト流し込み",
            "定義済みlayout_idから1〜50枚の新規デッキを構成",
            "白紙からタイトル・アジェンダ・KPI・比較・工程・タイムライン・編集可能グラフ等の視覚的なデッキを構成",
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
     Description("LibreChatにアップロード済みのPPTXを安全に検査し、編集候補を非同期で解析します。file_idが不明ならsourceFileIdを省略し、そのユーザーの最新PPTXを使ってください。")]
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
     Description("企業テンプレートの既存スライドにある名前付きテキストシェイプへ内容を流し込み、PPTXと全ページプレビューを作ります。")]
    public static Task<JobReceipt> PopulateTemplateAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("1始まりのスライド番号、設定文字列、解析で得たshapeNameまたはshapeIdからなるフィールド一覧。名前が重複する場合はshapeIdを指定します。")]
        IReadOnlyList<TemplateField> fields,
        CancellationToken cancellationToken,
        [Description("LibreChatへアップロード済みのテンプレートPPTXのfile_id。省略時はそのユーザーの最新PPTX。")]
        string sourceFileId = "latest") =>
        jobs.SubmitPopulateTemplateAsync(callerContext.GetRequired(), sourceFileId, fields, cancellationToken);

    [McpServerTool(Name = "pptx_create_deck", Destructive = true),
     Description("企業テンプレートのマスターと定義済みレイアウトを使い、1〜50枚の新規PPTXと全ページプレビューを作ります。先にpptx_analyzeとpptx_get_jobでlayout_idとplaceholderのshape_idを取得してください。slidesは必須で、完成版の全ページを1回の呼び出しに含めます。sourceFileIdだけで呼んではいけません。slidesが欠けた呼び出しにはinput_requiredを返すので、同じツールを全slides付きで直ちに再実行してください。各layout_id、shape_id、placeholder_indexは解析結果から一字も変更せずコピーし、推測したパスやIDを作らないでください。")]
    public static async Task<object> CreateDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("動作上必須。完成版の全1〜50ページからなる配列です。各要素はlayout_idとfields、各fieldはtextと解析結果のshape_id（任意でshape_nameまたはplaceholder_index）を使います。例: [{\"layout_id\":\"/ppt/slideLayouts/slideLayout1.xml\",\"fields\":[{\"text\":\"Title\",\"shape_id\":2}]}]。キーはこのsnake_caseを厳守し、全ページの内容を組み立て終えてから1回だけ呼びます。")]
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
     Description("アップロード済みテンプレートを使わず、検証済みの宣言型レイアウトから華やかで編集可能な16:9 PPTXを作ります。title/agenda/section/bullets/metrics/comparison/process/timeline/chart/quote/closingを内容に応じて使い分けてください。生成後はpptx_get_jobを待ち、必ずpptx_get_preview_imagesで全スライドを視覚確認し、問題があればdeckを修正して最大2回まで再生成してください。")]
    public static Task<JobReceipt> CreateVisualDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("資料タイトル、任意のテーマ、1〜50枚のスライドからなる宣言型仕様。各スライドは意味に合うkindを選び、1枚1メッセージに絞ります。theme.presetはmidnight/aurora/sunset/forest/minimal。任意のテーマ色はRRGGBBまたは#RRGGBB形式です。chartはPowerPoint上で編集可能です。")]
        VisualDeckSpec deck,
        CancellationToken cancellationToken) =>
        jobs.SubmitVisualDeckAsync(callerContext.GetRequired(), deck, cancellationToken);

    [McpServerTool(Name = "pptx_get_job", ReadOnly = true, Idempotent = true),
     Description("PowerPointジョブの状態、解析結果、短時間有効なプレビューURLとダウンロードURLを取得します。完了までpoll_after_secondsを目安に再実行してください。")]
    public static Task<JobView> GetJobAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("PowerPoint MCPが返したjob_id。")]
        string jobId,
        CancellationToken cancellationToken) =>
        jobs.GetAsync(callerContext.GetRequired(), jobId, cancellationToken);

    [McpServerTool(Name = "pptx_get_preview_images", ReadOnly = true, Idempotent = true),
     Description("成功済みジョブのスライド画像をClaude自身の視覚確認用に返します。全スライドを1〜4枚ずつ取得し、文字切れ・はみ出し・重なり・小さすぎる文字・余白・整列・コントラスト・情報階層・密度・バランス・資料全体の一貫性を確認してください。企業テンプレート資料の問題はpptx_refine_deckへ変更ページだけを渡して再生成し、白紙資料は仕様を修正して再生成します。最大2回で収束させ、このツールを呼ばずに視覚確認済みと述べてはいけません。")]
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
            Text = "Evaluate the returned slides for clipping, overflow, overlap, legibility, spacing, alignment, contrast, hierarchy, density, balance, and consistency. For a template-based deck, call pptx_refine_deck with only the changed slides when necessary; do not resend the complete deck to pptx_create_deck.",
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
}
