using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
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
            "導入環境に既定テンプレートが登録されていれば、新規Visual Deckへ自動適用する",
            "別テンプレートや既存資料を使う場合だけLibreChatへPPTXをアップロードする",
            "アップロードしたPPTXはpptx_analyzeでスライドと編集候補を取得する（file_idが会話に提示されない場合はsourceFileId=latestを使う）",
            "対象が曖昧ならスライド番号とshape_idをユーザーに選択してもらう",
            "企業テンプレートから新規作成する場合は、解析結果のlayout_idとshape_idをそのまま使い、全ページ分のslidesを1回のpptx_create_deckへ渡す",
            "企業テンプレートのマスター、ロゴ、フッターを保ちながら華やかな資料を作る場合はVisual Deckドラフトをpptx_finish_branded_visual_deckで生成する",
            "既存スライドを更新する場合はpptx_replace_textまたはpptx_populate_templateを使う",
            "新規資料はpptx_start_visual_deck、pptx_add_visual_slides_to_draft、pptx_finish_visual_deckの順で小分けに構成し、明示的に不要と言われない限り既定テンプレートを使う",
            "成功したVisual Deckへページを追加する場合はpptx_insert_visual_slidesへ追加分だけを渡し、既存ページを再送しない",
            "非同期ジョブ受領後はpptx_wait_for_jobでサーバー内待機し、短間隔の状態確認を繰り返さない",
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
            "成功したVisual DeckまたはブランドVisual Deckへの1〜49ページの追加（末尾追加または指定ページ直後への挿入）",
            "LibreOfficeによる全スライドPNGプレビュー",
            "プレビュー画像をClaudeへ返す自動視覚リフレクション",
            "署名付きURLによる成果物ダウンロード",
        },
        planned_after_template_spike = new[]
        {
            "SmartArtノードの追加・削除（種類変更は対象外）",
            "編集可能グラフのデータ・系列・凡例・軸・色の更新（種類変更は対象外）",
            "埋め込みExcelの値・式・範囲・行列更新",
            "任意の既存PPTXまたは厳密なプレースホルダー資料に対するスライド追加、およびスライド削除・並べ替え",
        },
    };

    [McpServerTool(Name = "pptx_analyze", ReadOnly = true, Idempotent = true),
     Description("登録済み既定テンプレートまたはLibreChatにアップロード済みのPPTXを安全に検査し、編集候補、企業テーマ色、日本語フォントを非同期で解析します。既定テンプレートはsourceFileId=default、添付のfile_idが不明ならlatestを使ってください。既定テンプレートは起動時解析キャッシュを再利用します。解析結果のthemeは白紙VisualDeckSpecのthemeへ移植できます。")]
    public static Task<JobReceipt> AnalyzeAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("既定テンプレートはdefault。添付PPTXはLibreChatのfile_id、会話にfile_idが提示されていなければlatest。ファイル名やパスは指定しません。省略時はlatest。")]
        string sourceFileId = "latest") =>
        jobs.SubmitAnalyzeAsync(callerContext.GetRequired(), sourceFileId, cancellationToken);

    [McpServerTool(Name = "pptx_render_preview", ReadOnly = true, Idempotent = true),
     Description("登録済み既定テンプレートまたはアップロード済みPPTXの全スライドをPNGへ変換する非同期ジョブを開始します。")]
    public static Task<JobReceipt> RenderPreviewAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("既定テンプレートはdefault。添付PPTXはLibreChatのfile_idまたはlatest。省略時はlatest。")]
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
     Description("企業テンプレートのマスターと定義済みレイアウトを使い、1〜50枚の新規PPTXと全ページプレビューを作ります。既定テンプレートにはsourceFileId=defaultを使い、厳密な既存プレースホルダー流し込みが必要な場合だけpptx_analyzeへsourceFileId=defaultを渡して起動時解析キャッシュからlayout_idとshape_idを取得します。アップロードした別テンプレートでも先にpptx_analyzeとpptx_wait_for_jobを実行してください。通常文はtext、箇条書き・番号付き手順はparagraphsを使い、記号や番号を本文へ手入力しません。slidesは必須で、完成版の全ページを1回の呼び出しに含めます。sourceFileIdだけで呼んではいけません。slidesが欠けた呼び出しにはinput_requiredを返すので、同じツールを全slides付きで直ちに再実行してください。各layout_id、shape_id、placeholder_indexは解析結果から一字も変更せずコピーし、推測したパスやIDを作らないでください。")]
    public static async Task<object> CreateDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Required, Description("動作上必須。完成版の全1〜50ページからなる配列です。各要素はlayout_idとfields、各fieldはtextまたはparagraphsのどちらか一方と解析結果のshape_id（任意でshape_nameまたはplaceholder_index）を使います。paragraphsの各項目はtext、kind=Plain/Bullet/Numbered、level=0〜4、任意のstart_atです。例: [{\"layout_id\":\"/ppt/slideLayouts/slideLayout1.xml\",\"fields\":[{\"paragraphs\":[{\"text\":\"現状把握\",\"kind\":\"Numbered\",\"level\":0}],\"shape_id\":2}]}]。キーはsnake_caseを厳守し、全ページを組み立て終えてから1回だけ呼びます。")]
        IReadOnlyList<DeckSlideSpec>? slides = null,
        [Description("既定テンプレートはdefault。添付した別テンプレートはLibreChatのfile_id、識別子が見えない場合はlatest。省略時はdefault。")]
        string sourceFileId = "default")
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
        [Required, Description("動作上必須。直前の成功したpptx_create_deckが返したjob_idをjobIdとして指定します。")]
        string jobId = "",
        [Required, Description("動作上必須。変更するページだけの配列です。各要素は1始まりのslide_numberと、そのページに残す全フィールドのfieldsをsnake_caseで指定します。fieldsの各要素はtextと元ページで使ったshape_idを含めます。")]
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

    [McpServerTool(Name = "pptx_start_visual_deck", Destructive = true),
     Description("新規Visual Deckの小さなサーバー側ドラフトを開始します。新規資料では必ず最初に1回だけ呼び、資料タイトル、最終ページ数、全体のthemeとdesignだけを指定します。この段階でスライド本文やPPTXは生成しません。返されたdraft_idへpptx_add_visual_slides_to_draftで1〜4ページずつ順番に追加してください。空引数で呼ばないでください。")]
    public static object StartVisualDeck(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        [Required, Description("資料全体のタイトル。1〜160文字。")]
        string title,
        [Required, Description("完成時の総ページ数。1〜50。後から変更しないため、構成を決めてから指定します。")]
        int expectedSlideCount,
        [Description("任意の全体テーマ。presetはmidnight/aurora/sunset/forest/minimal/ocean/berry/clay/cyber。")]
        VisualThemeSpec? theme = null,
        [Description("任意の資料サブジェクト。最大240文字。")]
        string? subject = null,
        [Description("言語コード。省略時ja-JP。")]
        string language = "ja-JP",
        [Description("任意の全体デザイン。style、density、motifを指定します。")]
        VisualDesignSpec? design = null)
    {
        try
        {
            return drafts.Begin(
                callerContext.GetRequired(),
                title,
                expectedSlideCount,
                theme,
                subject,
                language,
                design);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_start_visual_deck", exception);
        }
    }

    [McpServerTool(Name = "pptx_add_visual_slides_to_draft", Destructive = true),
     Description("新規Visual Deckのドラフトへ、順番どおりの完成済みVisualSlideSpecを1〜4ページだけ追加します。startSlideNumberは直前の応答が返したnext_slide_numberと一致させます。各ページは1枚1メッセージに絞り、title/agenda/section/statement/cards/metrics/comparison/process/timeline/matrix/funnel/roadmap/chart/dashboard/quote/closingを内容に応じて使い分けます。6枚以上の資料では全体で4種類以上の構図を計画し、単純な箇条書きは補助的に限定してください。既に受理されたページを再送しません。remaining_slide_countが0になるまでpptx_finish_*を呼びません。")]
    public static object AddVisualSlidesToDraft(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        [Required, Description("pptx_start_visual_deckが返したdraft_id。")]
        string draftId,
        [Required, Description("このバッチの先頭ページ番号。直前の応答にあるnext_slide_numberをそのまま使います。")]
        int startSlideNumber,
        [Required, Description("追加する連続した1〜4ページだけ。Metricsはmetricsを2〜6件、Dashboardはmetricsを2〜4件とchart、Cardsはcardsを3〜6件、Comparisonはpanelsを2〜3件、Process/Timeline/Funnel/Roadmapはstepsを3〜6件、Matrixはquadrantsを正確に4件指定します。metric/cardのtoneは意味語または#RRGGBB、iconは組み込み業務アイコンを使えます。")]
        IReadOnlyList<VisualSlideSpec> slides)
    {
        try
        {
            return drafts.AddSlides(
                callerContext.GetRequired(),
                draftId,
                startSlideNumber,
                slides);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_add_visual_slides_to_draft", exception);
        }
    }

    [McpServerTool(Name = "pptx_finish_visual_deck", Destructive = true),
     Description("全ページを追加済みのVisual DeckドラフトをPPTXとして生成します。pptx_start_visual_deckとpptx_add_visual_slides_to_draftを完了し、remaining_slide_count=0になった後にdraftId付きで1回だけ呼びます。導入環境に既定テンプレートがあれば自動的にマスター、ロゴ、フッター、テーマを適用します。ユーザーがテンプレート不要と明示した場合だけuseDefaultTemplate=falseにします。生成後はpptx_wait_for_jobで完了を待ち、全ページを視覚確認してください。")]
    public static Task<object> FinishVisualDeckAsync(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        JobService jobs,
        [Required, Description("全ページ追加後のVisual DeckドラフトID。")]
        string draftId,
        [Description("省略時true。導入環境の既定テンプレートを適用します。ユーザーがテンプレート不要と明示した場合だけfalse。")]
        bool useDefaultTemplate = true,
        CancellationToken cancellationToken = default) =>
        FinishVisualDraftAsync(
            "pptx_finish_visual_deck",
            callerContext,
            drafts,
            jobs,
            draftId,
            (caller, deck, token) => jobs.SubmitVisualDeckAsync(caller, deck, useDefaultTemplate, token),
            cancellationToken);

    [McpServerTool(Name = "pptx_finish_branded_visual_deck", Destructive = true),
     Description("全ページを追加済みのVisual Deckドラフトを、既定またはアップロード済み企業テンプレートのマスター、ロゴ、フッター、ページ設定へ合成します。通常の新規資料はpptx_finish_visual_deckが既定テンプレートを自動適用するため、主に添付した別テンプレートを明示的に使う場合に呼びます。事前に添付を解析し、remaining_slide_count=0になったドラフトIDを指定してください。templateLayoutIdは通常autoです。")]
    public static Task<object> FinishBrandedVisualDeckAsync(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        JobService jobs,
        [Required, Description("全ページ追加後のVisual DeckドラフトID。")]
        string draftId,
        [Description("通常はauto。明示する場合はpptx_analyzeが返したプレースホルダー0個の白紙layout_idを一字も変更せず指定します。")]
        string templateLayoutId = "auto",
        [Description("既定テンプレートはdefault。添付した別テンプレートはLibreChatのfile_id、識別子が見えない場合はlatest。省略時はdefault。")]
        string sourceFileId = "default",
        CancellationToken cancellationToken = default) =>
        FinishVisualDraftAsync(
            "pptx_finish_branded_visual_deck",
            callerContext,
            drafts,
            jobs,
            draftId,
            (caller, deck, token) => jobs.SubmitBrandedVisualDeckAsync(
                caller,
                sourceFileId,
                deck,
                templateLayoutId,
                token),
            cancellationToken);

    [McpServerTool(Name = "pptx_insert_visual_slides", Destructive = true),
     Description("成功したVisual DeckまたはブランドVisual Deckへ新しいページを追加します。ユーザーが追加・挿入・末尾へ追加を依頼した場合は、新しいVisual Deckドラフトで資料全体を作り直さず、このツールへ追加分の完全なVisualSlideSpecだけを渡してください。jobIdは通常latest、afterSlideNumberは省略時に末尾追加、0なら先頭追加、正の値ならその1始まりページの直後へ挿入します。既存の資料タイトル、テーマ、design、全ページ、企業テンプレートと選択レイアウトはサーバー側で継承します。通常はサーバー内で完了まで待って最終statusと成果物を返すため、Succeededなら状態確認ツールを呼ばず全ページを視覚確認してください。30秒以内に完了せずqueuedを返した場合だけjob_idをpptx_wait_for_jobで待ちます。")]
    public static async Task<object> InsertVisualSlidesAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Required, Description("動作上必須。追加する1ページ以上のVisualSlideSpecだけを含む配列。既存ページを含めず、資料全体を再送しません。")]
        IReadOnlyList<VisualSlideSpec>? slides = null,
        [Description("省略時は末尾へ追加。0は先頭、正の値はその1始まりページの直後へ挿入します。既存ページ数を超えてはいけません。")]
        int? afterSlideNumber = null,
        [Description("通常はlatest。明示する場合は直前に成功したVisual Deckジョブのjob_id。latestは同じ会話の最新成功ジョブを安全に選びます。")]
        string jobId = "latest")
    {
        var requestedSlides = slides ?? [];
        if (requestedSlides.Count == 0)
        {
            return new ToolInputRequest(
                "input_required",
                "pptx_insert_visual_slides",
                ["slides"],
                "Call pptx_insert_visual_slides again with only the new slides. Do not resend existing slides or call a create tool.");
        }

        try
        {
            var caller = callerContext.GetRequired();
            var resolvedJobId = string.IsNullOrWhiteSpace(jobId)
                || string.Equals(jobId, "latest", StringComparison.OrdinalIgnoreCase)
                ? await jobs.GetLatestSuccessfulVisualJobIdAsync(caller, cancellationToken).ConfigureAwait(false)
                : jobId;
            var receipt = await jobs.SubmitInsertVisualSlidesAsync(
                caller,
                resolvedJobId,
                requestedSlides,
                afterSlideNumber,
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
            return CreateValidationError("pptx_insert_visual_slides", exception);
        }
    }

    [McpServerTool(Name = "pptx_refine_visual_deck", Destructive = true),
     Description("成功したpptx_finish_visual_deckまたはpptx_finish_branded_visual_deckジョブの仕様を再利用し、視覚確認で問題があったページだけを差し替えて再生成します。ブランドVisual Deckでは元の企業テンプレートと選択レイアウトも自動的に維持します。資料全体を再送せず、変更ページの完全なVisualSlideSpecだけをrevisionsへ指定してください。jobIdまたはrevisionsが欠けた場合はinput_requiredを返します。")]
    public static async Task<object> RefineVisualDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Required, Description("動作上必須。直前の成功したpptx_finish_visual_deckまたはpptx_finish_branded_visual_deckが返したjob_id。")]
        string jobId = "",
        [Required, Description("動作上必須。変更ページだけの配列。各要素は1始まりのslide_numberと、差し替え後の完全なslideを含めます。")]
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
     Description("成功したVisual DeckまたはブランドVisual Deckの既存ページを1枚だけ差し替えます。ページ追加には使わずpptx_insert_visual_slidesを使ってください。Bedrock/Claudeでの自動視覚リフレクションでは、大きなrevisions配列を作るpptx_refine_visual_deckよりこのツールを優先してください。revisionにslide_numberと差し替え後の完全なslideを必ず含めます。jobIdは通常latestを使い、会話内の最新成功Visual Deckを自動選択します。通常はサーバー内で完了まで待って最終statusと成果物を返すため、Succeededなら状態確認ツールを呼ばず次のページへ進んでください。30秒以内に完了せずqueuedを返した場合だけjob_idをpptx_wait_for_jobで待ちます。複数ページを1ページずつ直すと修正が累積します。全ページを最大2巡までに収束させてください。")]
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
     Description("PowerPointジョブの現在状態を待たずに1回だけ取得します。障害復旧や即時確認用です。待機中・実行中ジョブの通常フローでは、短間隔でこのツールを反復せずpptx_wait_for_jobを使ってください。jobId=latestでは同じ利用者・会話の直近ジョブを安全に選びます。")]
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

    [McpServerTool(Name = "pptx_wait_for_job", ReadOnly = true, Idempotent = true),
     Description("PowerPointジョブをMCPサーバー内で最大45秒待ち、終端状態になれば解析結果、プレビューURL、ダウンロードURLを返します。非同期ジョブ受領後はpptx_get_jobの短間隔ポーリングをせず、このツールを1回呼んでください。指定時間後もRunningまたはQueuedなら、同じjobIdでもう一度だけ待機できます。")]
    public static async Task<object> WaitForJobAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("PowerPoint MCPが返したjob_id、または同じ会話の直近ジョブを選ぶlatest。省略時はlatest。")] string jobId = "latest",
        [Description("サーバー内で待つ秒数。省略時45秒、1〜50秒。LibreChatのMCPタイムアウトより短くします。")] int waitSeconds = 45)
    {
        if (waitSeconds is < 1 or > 50)
        {
            return new ToolValidationError(
                "invalid_input",
                "pptx_wait_for_job",
                "wait_seconds_invalid",
                "waitSeconds must be between 1 and 50.",
                "Call pptx_wait_for_job again with waitSeconds between 1 and 50.");
        }

        try
        {
            return await jobs.WaitAsync(
                callerContext.GetRequired(),
                jobId,
                TimeSpan.FromSeconds(waitSeconds),
                cancellationToken).ConfigureAwait(false);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_wait_for_job", exception);
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

    private static async Task<object> FinishVisualDraftAsync(
        string toolName,
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        JobService jobs,
        string draftId,
        Func<CallerContext, VisualDeckSpec, CancellationToken, Task<JobReceipt>> submit,
        CancellationToken cancellationToken)
    {
        var caller = callerContext.GetRequired();
        var acquired = false;
        try
        {
            var submission = drafts.AcquireForSubmission(caller, draftId);
            if (submission.ExistingJobId is not null)
            {
                return await jobs.GetAsync(caller, submission.ExistingJobId, cancellationToken).ConfigureAwait(false);
            }

            acquired = true;
            var receipt = await submit(caller, submission.Deck!, cancellationToken).ConfigureAwait(false);
            drafts.MarkSubmitted(caller, draftId, receipt.JobId);
            return receipt;
        }
        catch (PptxValidationException exception)
        {
            if (acquired)
            {
                drafts.ReleaseSubmission(caller, draftId);
            }

            return CreateValidationError(toolName, exception);
        }
        catch
        {
            if (acquired)
            {
                drafts.ReleaseSubmission(caller, draftId);
            }

            throw;
        }
    }

    private static ToolValidationError CreateValidationError(string tool, PptxValidationException exception) =>
        new(
            "invalid_input",
            tool,
            exception.Code,
            exception.Message,
            $"Correct the field named in message and call {tool} again. Do not repeat the same invalid input.");
}
