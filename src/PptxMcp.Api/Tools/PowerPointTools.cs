using System.ComponentModel;
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
            "pptx_analyzeでスライドと編集候補を取得する",
            "対象が曖昧ならスライド番号とshape_idをユーザーに選択してもらう",
            "pptx_create_deck、pptx_replace_text、pptx_populate_templateのいずれかを実行する",
            "pptx_get_jobで完了を確認し、全ページのプレビューとPPTXを提示する",
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
            "LibreOfficeによる全スライドPNGプレビュー",
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
     Description("LibreChatにアップロード済みのPPTXを安全に検査し、編集候補を非同期で解析します。ローカルパスではなくfile_idを指定してください。")]
    public static Task<JobReceipt> AnalyzeAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("LibreChatのアップロード済みPPTXのfile_id。ファイル名やパスではありません。")]
        string sourceFileId,
        CancellationToken cancellationToken) =>
        jobs.SubmitAnalyzeAsync(callerContext.GetRequired(), sourceFileId, cancellationToken);

    [McpServerTool(Name = "pptx_render_preview", ReadOnly = true, Idempotent = true),
     Description("アップロード済みPPTXの全スライドをPNGへ変換する非同期ジョブを開始します。")]
    public static Task<JobReceipt> RenderPreviewAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("LibreChatのアップロード済みPPTXのfile_id。")]
        string sourceFileId,
        CancellationToken cancellationToken) =>
        jobs.SubmitRenderAsync(callerContext.GetRequired(), sourceFileId, cancellationToken);

    [McpServerTool(Name = "pptx_replace_text", Destructive = true),
     Description("PPTXの文字列を置換し、更新版PPTXと全ページプレビューを作る非同期ジョブを開始します。曖昧な対象は先にpptx_analyzeで特定してください。")]
    public static Task<JobReceipt> ReplaceTextAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("LibreChatのアップロード済みPPTXのfile_id。")]
        string sourceFileId,
        [Description("検索文字、置換文字、任意の1始まりスライド番号、任意のshapeNameとshapeIdからなる置換指示。名前が重複する場合はshapeIdも指定します。")]
        IReadOnlyList<TextReplacement> replacements,
        CancellationToken cancellationToken) =>
        jobs.SubmitReplaceTextAsync(callerContext.GetRequired(), sourceFileId, replacements, cancellationToken);

    [McpServerTool(Name = "pptx_populate_template", Destructive = true),
     Description("企業テンプレートの既存スライドにある名前付きテキストシェイプへ内容を流し込み、PPTXと全ページプレビューを作ります。")]
    public static Task<JobReceipt> PopulateTemplateAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("LibreChatへアップロード済みのテンプレートPPTXのfile_id。")]
        string sourceFileId,
        [Description("1始まりのスライド番号、設定文字列、解析で得たshapeNameまたはshapeIdからなるフィールド一覧。名前が重複する場合はshapeIdを指定します。")]
        IReadOnlyList<TemplateField> fields,
        CancellationToken cancellationToken) =>
        jobs.SubmitPopulateTemplateAsync(callerContext.GetRequired(), sourceFileId, fields, cancellationToken);

    [McpServerTool(Name = "pptx_create_deck", Destructive = true),
     Description("企業テンプレートのマスターと定義済みレイアウトを使い、1〜50枚の新規PPTXと全ページプレビューを作ります。先にpptx_analyzeでlayout_idとplaceholderのshape_idを取得してください。")]
    public static Task<JobReceipt> CreateDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("LibreChatへアップロード済みの企業テンプレートPPTXのfile_id。")]
        string sourceFileId,
        [Description("layoutIdとfieldsからなるスライド一覧。fieldsはtextと解析結果のshapeId、任意でshapeNameまたはplaceholderIndexを指定します。")]
        IReadOnlyList<DeckSlideSpec> slides,
        CancellationToken cancellationToken) =>
        jobs.SubmitCreateDeckAsync(callerContext.GetRequired(), sourceFileId, slides, cancellationToken);

    [McpServerTool(Name = "pptx_get_job", ReadOnly = true, Idempotent = true),
     Description("PowerPointジョブの状態、解析結果、短時間有効なプレビューURLとダウンロードURLを取得します。完了までpoll_after_secondsを目安に再実行してください。")]
    public static Task<JobView> GetJobAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("PowerPoint MCPが返したjob_id。")]
        string jobId,
        CancellationToken cancellationToken) =>
        jobs.GetAsync(callerContext.GetRequired(), jobId, cancellationToken);

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
