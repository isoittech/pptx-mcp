using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Collections.Concurrent;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Jobs;
using PptxMcp.Security;
using PptxMcp.Storage;

namespace PptxMcp.Tools;

[McpServerToolType]
public sealed class PowerPointTools
{
    private static readonly ConcurrentDictionary<string, byte> DeliveredAnalysisResults = new();

    [McpServerTool(Name = "pptx_prepare_visual_objects", Destructive = true),
     Description("スライド理解を助ける矢印・枠・吹き出し・括弧・リング・リボンを、PowerPointで編集可能なネイティブ図形として一括準備します。ストーリーとAsset Planを組み立てた後、Design Briefをvalidate/prepareする前に必要な場合だけ1回呼び、1〜8件をまとめます。recipe=directionalCue|growthPath|focusCorners|annotationPin|sectionRule|cycleCueでは、安全な意味anchorへ控えめな複合図形を配置します。annotationPinは折れ線Chart上の実在点を、1始まりのcategory／series番号で指定します。autoは従来の単一図形です。座標、任意色、SVG/XML、URL、path、コードは受け取りません。返ったopaque asset_idを同じslideのassetPlan.visual_object_asset_idsとslide.visualObjectsへコピーします。1ページ最大3個、会話全体最大24個です。複合recipeは原則subtleまたはstandardとし、strongはvisualPurpose=emphasisのfocusCornersだけに使えます。装飾目的だけでは使わず、方向・成長・焦点・注釈・区切り・循環のいずれかを説明するときだけ使います。")]
    public static CallToolResult PrepareVisualObjects(
        CallerContextAccessor callerContext,
        VisualObjectAssetRepository visualObjects,
        [Required, Description("1〜8件の意味仕様。slideNumber、visualPurpose、archetype、style、emphasis、orientation、placementRole、paletteRole、必要ならrecipeを指定します。recipeの正規tuple: directionalCue=direction+arrow+headerAccent/contentConnector、growthPath=growth+arrow+right/up+contentConnector/chartAnnotation、focusCorners=emphasis/grouping+frame+focusFrame、annotationPin=annotation+callout+chartAnnotation+label+anchorCategoryOrdinal(1-12)+anchorSeriesOrdinal(1-4)、sectionRule=emphasis/annotation+ribbon+headerAccent/sectionDivider、cycleCue=cycle+curvedArrow/ring+clockwise+contentConnector/backgroundMotif。annotationPinは同じページの折れ線Chartにある実在点だけを1始まりで指定し、他recipeはanchor番号を省略します。複合recipeのstrongはemphasis+focusCornersだけで、他はsubtle/standardです。生の座標・色・SVG/XML・URL・pathは指定できません。")]
        IReadOnlyList<VisualObjectBrief> objects)
    {
        try
        {
            return VisualObjectPreviewResource.Create(
                visualObjects.Prepare(callerContext.GetRequired(), objects));
        }
        catch (PptxValidationException exception)
        {
            var error = CreateValidationError("pptx_prepare_visual_objects", exception);
            return new CallToolResult
            {
                Content = [new TextContentBlock { Text = JsonSerializer.Serialize(error) }],
                IsError = true,
            };
        }
    }

    [McpServerTool(Name = "pptx_register_uploaded_image_asset", Destructive = true),
     Description("LibreChatへ添付済みのJPEG/PNGを、現在の利用者・会話だけで使える期限付き画像assetへ登録します。URLやローカルpathは受け取りません。原本を検証・回転・sRGB化・縮小し、metadataを除いたPNGへ無害化してopaque asset_idを返します。画像が必要なAsset Planを確定する前に呼び、返ったasset_idを同じslideのassetPlan.asset_idとMedia.media.assetIdへコピーしてください。sourceFileIdが会話に提示されない場合はlatestを使えます。利用者が提供権限を持つ画像だけを登録し、altTextには画像の内容、attributionRefには任意のopaque出典記録IDだけを指定します。")]
    public static async Task<object> RegisterUploadedImageAssetAsync(
        CallerContextAccessor callerContext,
        ImageAssetRepository imageAssets,
        [Required, Description("画像を見られない利用者にも意味が伝わる、1〜240文字の簡潔な代替テキスト。")] string altText,
        [Description("LibreChat添付のopaque file_id。省略時は最新のJPEG/PNG添付を選ぶlatest。URL、path、ファイル名は不可。")] string sourceFileId = "latest",
        [Description("任意のopaque出典記録ID。URL、path、出典本文は不可。")] string? attributionRef = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            return await imageAssets.RegisterUserUploadAsync(
                callerContext.GetRequired(),
                sourceFileId,
                altText,
                attributionRef,
                cancellationToken).ConfigureAwait(false);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_register_uploaded_image_asset", exception);
        }
    }

    [McpServerTool(Name = "pptx_get_capabilities", ReadOnly = true, Idempotent = true),
     Description("PowerPoint MCPの対応範囲、制約、推奨ワークフローを返します。")]
    public static object GetCapabilities() => new
    {
        workflow = new[]
        {
            "導入環境に既定テンプレートが登録されていれば、新規Visual Deckへ自動適用する",
            "Brand Profileが外部登録されている場合はpptx_get_design_catalogで用途・密度・方向に合うrecipeとsample要約を選ぶ",
            "Design Brief必須の導入環境では全ページのAsset Planを確定し、pptx_validate_design_briefが返したbrief_idをstartより前に取得する",
            "デザイン方向の選択が結果へ大きく影響するときだけpptx_prepare_design_briefでカードを表示し、ユーザー選択後にpptx_apply_design_brief_actionが返すbrief_idを使う",
            "別テンプレートや既存資料を使う場合だけLibreChatへPPTXをアップロードする",
            "ユーザー提供画像を使う場合はJPEG/PNGをLibreChatへ添付し、pptx_register_uploaded_image_assetで無害化済みopaque asset_idへ登録してからDesign Briefを確定する",
            "意味のある矢印・枠・吹き出し等が必要な場合はpptx_prepare_visual_objectsで最大8件を一括準備し、Asset Planへ束縛してからDesign Briefを確定する",
            "アップロードしたPPTXはpptx_analyzeでスライドと編集候補を取得する（file_idが会話に提示されない場合はsourceFileId=latestを使う）",
            "対象が曖昧ならスライド番号とshape_idをユーザーに選択してもらう",
            "企業テンプレートから新規作成する場合は、解析結果のlayout_idとshape_idをそのまま使い、全ページ分のslidesを1回のpptx_create_deckへ渡す",
            "企業テンプレートのマスター、ロゴ、フッターを保ちながら華やかな資料を作る場合はVisual Deckドラフトをpptx_finish_branded_visual_deckで生成する",
            "既存スライドを更新する場合はpptx_replace_textまたはpptx_populate_templateを使う",
            "新規資料はpptx_start_visual_deckでテンプレート・theme・designを固定し、pptx_add_visual_slides_to_draft、pptx_finish_visual_deckの順で小分けに構成する",
            "pptx_add_visual_slides_to_draftのstartSlideNumberは省略し、サーバー側で現在の末尾から自動計算する",
            "文字量が多い説明はstructured_brief、複数案を評価軸で比べる場合はscorecardを使い、小さい文字へ一括縮小しない",
            "五線譜とウクレレTABを編集可能なPowerPoint図形で併記する場合はmusicScoreを使い、音高、音価、弦、フレット、指番号を指定する",
            "成功したVisual Deckへページを追加する場合はpptx_insert_visual_slidesへ追加分だけを渡し、既存ページを再送しない",
            "非同期ジョブ受領後はpptx_wait_for_jobでサーバー内待機し、短間隔の状態確認を繰り返さない",
            "pptx_get_preview_imagesで全ページを実際に見て、企業テンプレート資料はpptx_refine_deck、白紙・ブランドVisual Deckはpptx_refine_visual_slideへ問題ページを1枚ずつ渡してサーバー強制の最大3巡で収束させる",
            "Visual Deck生成が一度成功したら全体生成を再開始せず、問題ページの差分修正または追加ページだけの挿入を使う",
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
            "通常シェイプ・編集可能なPowerPoint表・SmartArt内部テキストの置換",
            "名前付きシェイプを持つ企業テンプレートへのテキスト流し込み",
            "定義済みlayout_idから1〜50枚の新規デッキを構成",
            "白紙からKPI・カード・比較・構造化ブリーフ・編集可能スコアカード・工程・タイムライン・マトリクス・ファネル・ロードマップ・ダッシュボード・編集可能グラフ等の視覚的なデッキを構成",
            "外部読み取り専用Brand Profileのimmutable version/hash、用途別recipe、情報密度別sample要約、禁止・要確認ルールをcompact catalogとして取得",
            "Design Briefと全ページAsset Planを検証し、利用者・会話・Brand Profileへ束縛した期限付きopaque brief_idを発行",
            "条件付きDesign Briefカード、最大3つのサーバー検証済み候補、完成サンプルcarousel、1回限りの利用者選択",
            "会話へ束縛したユーザー提供JPEG/PNGの検証・metadata除去・sRGB PNG化と、Media/splitへの実画像埋め込み、crop、代替テキスト、出典表示",
            "会話へ束縛した意味仕様から、矢印・曲線矢印・枠・吹き出し・括弧・リング・リボンと、方向線・成長線・コーナー強調枠・注釈ピン・セクション罫線・循環キューの複合レシピを編集可能なネイティブ図形として生成",
            "tree・flow・cycle・concentric・networkの編集可能NativeDiagramと、loop・stepped・pyramidの実装済み構図variant",
            "未選択カードが表示できない場合のcaller-boundな安全なcancelとdirect validationへの復帰",
            "五線、音符、休符、小節線、ウクレレTAB、色分けした指番号をPowerPointネイティブの線・図形・テキストとして生成",
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

    [McpServerTool(Name = "pptx_get_design_catalog", ReadOnly = true, Idempotent = true),
     Description("導入環境が外部登録したBrand Profileを取得します。新規資料では正確に2回だけ使います。無引数の1回目はcompactなprofile一覧と選択可能なstyle_directionsを返します。2回目はprofileIdと選んだstyleDirectionIdを必須のつもりで指定し、一致する色・文章・視覚ルール、意味レイアウトrecipe、完成サンプル要約を最終取得します。単一用途だけpurpose／densityで絞れますが、複数用途の資料はstyleDirectionIdだけで全用途recipeを取得します。2回目の後は再呼出しません。会社固有値をOSSへ埋め込まず、返されたID、version、content_hashだけを後続ツールへコピーしてください。")]
    public static object GetDesignCatalog(
        BrandProfileCatalog catalog,
        [Description("任意。無引数一覧から選んだ正確なprofile ID。URLやパスではありません。")]
        string? profileId = null,
        [Description("任意。profileIdと一緒に指定し、catalog内の用途IDでrecipeを絞ります。例: cover、kpi、comparison。自然文やパスではありません。")]
        string? purpose = null,
        [Description("任意。profileIdと一緒に指定し、airy、balanced、detailedのいずれかでrecipeを絞ります。")]
        string? density = null,
        [Description("任意。profileIdと一緒に指定し、選択profileの正確なstyle direction IDでrecipeを絞ります。")]
        string? styleDirectionId = null)
    {
        try
        {
            return catalog.Query(profileId, purpose, density, styleDirectionId);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_get_design_catalog", exception);
        }
    }

    [McpServerTool(Name = "pptx_validate_design_brief"),
     Description("PPTX生成前にDesign Briefと全ページのAsset Planを検証し、設定された有効期間内のopaqueなbrief_idを発行します。未解決質問、要確認の仮定、profile recipeとの不一致、権利未確認素材の採用、利用不能な画像計画は拒否します。ユーザー提供画像を使うページは先にpptx_register_uploaded_image_assetを呼び、acquisition=userUpload, status=ready, license_status=userProvided, fallback=noneと検証済みasset_id、crop_intent、text_safe_area=left|right、画像必須recipeを指定します。素材を使わないページはpreferred_medium=none, acquisition=none, fallback=none, status=omitted, license_status=notRequiredとし、noAssetLayoutを指定しません。approvedLibraryまたは未登録userUploadは画像なしrecipeへfallbackSelectedで切り替えます。このツールはPPTXを生成しません。")]
    public static object ValidateDesignBrief(
        CallerContextAccessor callerContext,
        DesignBriefService designBriefs,
        [Required, Description("audience、purpose、delivery_mode、desired_tone、density、brand_profile{id,version,content_hash}、style_direction_id、visual_strategy、source_policy、expected_slide_count、assumptions、空のquestions_for_userからなる確定済みbrief。")]
        DesignBriefSpec brief,
        [Required, Description("完成版の全ページに1件ずつ、slide_number順で指定する計画。各項目はcatalogのpurpose/recipe_id、visual_purpose、preferred_medium、acquisition、fallback、status、license_statusを持ちます。ready userUploadは登録済みasset_id、crop_intent、text_safe_area=left|rightを持たせます。素材なしの正確な組合せはpreferred_medium=none, acquisition=none, fallback=none, status=omitted, license_status=notRequiredで、asset metadataを省略し、required_asset_rolesが空のrecipeを使います。noAssetLayoutはacquisition=noneでは禁止です。URL、パス、画像バイナリは渡しません。")]
        IReadOnlyList<AssetPlanItem> assetPlan)
    {
        try
        {
            return designBriefs.Validate(callerContext.GetRequired(), brief, assetPlan);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_validate_design_brief", exception);
        }
    }

    [McpServerTool(Name = "pptx_prepare_design_brief"),
     Description("デザイン方向が結果へ大きく影響し、実行可能で見た目の異なる案が2件以上ある場合だけDesign Briefカードを準備します。推奨brief/Asset Planは完全形、別案はstyle/density/全slideのrecipe IDだけ（最大2件）、画像を使わない別構成はphoto計画slideのcanonical overrideだけを渡します。選択肢は合計2〜3件で、画像なし別構成があれば別styleは最大1件です。UI Resource返却後はターンを終了し、固定intent pptx.designBrief.selectを待ちます。pending中はvalidate/start禁止です。1案だけならpptx_validate_design_briefを使います。")]
    public static CallToolResult PrepareDesignBrief(
        CallerContextAccessor callerContext,
        DesignBriefService designBriefs,
        [Required, Description("推奨案の確定済みDesign Brief。questions_for_userは空で、profile ID/version/content_hashをcatalogから正確にコピーします。")]
        DesignBriefSpec brief,
        [Required, Description("推奨案の全ページAsset Plan。完成版のslide順に正確に1件ずつ指定します。")]
        IReadOnlyList<AssetPlanItem> assetPlan,
        [Description("任意の別style案、最大2件。各項目は同じprofile内のstyle_direction_id、density、slide順のrecipe_idsだけを指定し、共通briefやAsset Plan全文を重複させません。")]
        IReadOnlyList<DesignBriefStyleAlternative>? alternatives = null,
        [Description("任意の画像を使わない別構成。推奨案でpreferred_medium=photoの全slideだけを1件ずつ指定し、異なるrecipeとnativeDrawまたはcanonical noneへ置換します。第2段階は推奨案も外部画像を挿入しないため、単なる写真有無ではなく生成構造が変わるrecipeが必要です。")]
        IReadOnlyList<DesignBriefNoPhotoOverride>? noPhotoOverrides = null,
        [Description("通常false。未選択カードを置換するようユーザーが明示した場合だけtrue。古いカードの選択は無効になります。")]
        bool replacePendingChoice = false)
    {
        try
        {
            var card = designBriefs.Prepare(
                callerContext.GetRequired(),
                brief,
                assetPlan,
                alternatives,
                noPhotoOverrides,
                replacePendingChoice);
            try
            {
                return DesignBriefCardResource.Create(card);
            }
            catch
            {
                designBriefs.DiscardPrepared(callerContext.GetRequired(), card.ChoiceSessionId);
                throw;
            }
        }
        catch (PptxValidationException exception)
        {
            return DesignBriefCardResource.CreateError("pptx_prepare_design_brief", exception);
        }
    }

    [McpServerTool(Name = "pptx_apply_design_brief_action"),
     Description("Design Briefカードの固定intent pptx.designBrief.selectが返した不透明なchoiceSessionIdとoptionIdだけを適用します。表示文、style名、action名、brief_idを入力として信用しません。サーバーが利用者、会話、有効期限、Brand Profile version/hash、許可済み候補を照合し、選択済み候補だけをstart可能にします。同じoptionの二重送信はstart前に限り冪等で、別optionのreplay、改ざん、期限切れ、別利用者・別会話を拒否します。成功時のbrief_idをpptx_start_visual_deckへ渡してください。")]
    public static object ApplyDesignBriefAction(
        CallerContextAccessor callerContext,
        DesignBriefService designBriefs,
        [Required, Description("UI Resource intentが返した32桁のopaque choice session ID。推測や再構成をしません。")]
        string choiceSessionId,
        [Required, Description("同じintentが返した32桁のopaque option ID。style名やaction文字列へ置換しません。")]
        string optionId)
    {
        try
        {
            return designBriefs.ApplyCardSelection(
                callerContext.GetRequired(),
                choiceSessionId,
                optionId);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_apply_design_brief_action", exception);
        }
    }

    [McpServerTool(Name = "pptx_cancel_design_brief_selection", Destructive = true),
     Description("未選択のDesign Briefカードがホストで表示できない、またはユーザーがカードを使わず安全な既定案へ進むと明示した場合だけ、その利用者・会話のpending選択を破棄します。引数や選択IDは受け取りません。apply済み、start予約中、start済みの選択は取り消せません。成功後はpptx_validate_design_briefで安全な推奨案を新規確定するか、見た目の異なる2案以上で新しいカードを準備できます。")]
    public static object CancelDesignBriefSelection(
        CallerContextAccessor callerContext,
        DesignBriefService designBriefs)
    {
        try
        {
            return designBriefs.CancelPendingSelection(callerContext.GetRequired());
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_cancel_design_brief_selection", exception);
        }
    }

    [McpServerTool(Name = "pptx_analyze", ReadOnly = true, Idempotent = true),
     Description("登録済み既定テンプレートまたはLibreChatにアップロード済みのPPTXを安全に検査し、編集候補、企業テーマ色、日本語フォントを非同期で解析します。通常の翻訳・文字編集は既定のincludeLayouts=falseで、全スライドの構造化テキストをコンパクトに取得します。同じ利用者メッセージで同じsourceFileIdとincludeLayoutsを再送した場合は新しい解析を作らず、最初のjob_idを返します。pptx_create_deckで既存プレースホルダーへ厳密に流し込む場合だけincludeLayouts=trueを指定します。既定テンプレートはsourceFileId=default、添付のfile_idが不明ならlatestを使ってください。既定テンプレートは起動時解析キャッシュを再利用します。解析結果のthemeは白紙VisualDeckSpecのthemeへ移植できます。")]
    public static Task<object> AnalyzeAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("既定テンプレートはdefault。添付PPTXはLibreChatのfile_id、会話にfile_idが提示されていなければlatest。ファイル名やパスは指定しません。省略時はlatest。")]
        string sourceFileId = "latest",
        [Description("通常の翻訳・文字編集はfalse。pptx_create_deck用に既存プレースホルダーのlayout_id一覧が必要な場合だけtrue。省略時false。")]
        bool includeLayouts = false) =>
        ExecuteJobSubmissionAsync(
            "pptx_analyze",
            () => jobs.SubmitAnalyzeAsync(
                callerContext.GetRequired(),
                sourceFileId,
                cancellationToken,
                includeLayouts));

    [McpServerTool(Name = "pptx_render_preview", ReadOnly = true, Idempotent = true),
     Description("登録済み既定テンプレートまたはアップロード済みPPTXの全スライドをPNGへ変換する非同期ジョブを開始します。文字だけの翻訳・置換では、元資料の文字確認に使わず、pptx_analyzeの構造化テキストで置換した後の見た目確認に使ってください。同じユーザーメッセージですでに文字解析した元資料は例外なく拒否します。ユーザーが元資料の画像確認を明示した場合は、同じメッセージでpptx_analyzeを先に呼ばず、このツールを直接使います。")]
    public static Task<object> RenderPreviewAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Description("既定テンプレートはdefault。添付PPTXはLibreChatのfile_idまたはlatest。省略時はlatest。")]
        string sourceFileId = "latest") =>
        ExecuteJobSubmissionAsync(
            "pptx_render_preview",
            () => jobs.SubmitRenderAsync(
                callerContext.GetRequired(),
                sourceFileId,
                cancellationToken));

    [McpServerTool(Name = "pptx_replace_text", Destructive = true),
     Description("PPTXの文字列を最大20件ずつ置換する非同期ジョブです。20件を超える翻訳は複数バッチに分けます。最初はpreviousJobIdを省略し、残りがあればisFinalBatch=falseにします。成功後は返されたjob_idを次のpreviousJobIdへ正確に渡すと、直前までの置換を保持して続行します。最後だけisFinalBatch=trueにすると全ページプレビューを作ります。翻訳ではpptx_analyze(includeLayouts=false)の文字を使い、元資料の画像から転記しません。")]
    public static async Task<object> ReplaceTextAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("1〜20件の置換指示をJSON配列そのものとして渡します。JSON文字列へ変換してはいけません。各要素は検索文字、置換文字、任意の1始まりスライド番号、任意のshapeNameとshapeIdです。解析結果のslide_numberとshape_idを使います。")]
        IReadOnlyList<TextReplacement> replacements,
        CancellationToken cancellationToken,
        [Description("LibreChatの現在のメッセージへ添付したPPTXのfile_id。省略時は今回の添付内で最新のPPTX。")]
        string sourceFileId = "latest",
        [Description("2バッチ目以降だけ、直前に成功したpptx_replace_textのjob_idを指定します。ファイルIDやlatestではありません。")]
        string? previousJobId = null,
        [Description("後続バッチがある間はfalse。最後のバッチだけtrueにして全ページプレビューを生成します。省略時true。")]
        bool isFinalBatch = true)
    {
        if (replacements.Count is < 1 or > 20)
        {
            return new ToolValidationError(
                "invalid_input",
                "pptx_replace_text",
                "replacement_batch_size_invalid",
                "Each text replacement batch must contain between 1 and 20 entries.",
                "Split the complete replacement list into batches of at most 20. Submit the first batch with isFinalBatch=false when more remain, wait for success, then pass its exact job_id as previousJobId on the next batch. Set isFinalBatch=true only on the last batch.");
        }

        return await ExecuteJobSubmissionAsync(
            "pptx_replace_text",
            () => jobs.SubmitReplaceTextAsync(
                callerContext.GetRequired(),
                sourceFileId,
                replacements,
                cancellationToken,
                previousJobId,
                isFinalBatch)).ConfigureAwait(false);
    }

    [McpServerTool(Name = "pptx_populate_template", Destructive = true),
     Description("企業テンプレートの既存スライドにある名前付きテキストシェイプへ内容を流し込み、PPTXと全ページプレビューを作ります。通常文はtext、箇条書き・番号付き手順はparagraphsを使い、記号や番号を本文へ手入力しないでください。")]
    public static Task<object> PopulateTemplateAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("1始まりのスライド番号、解析で得たshapeNameまたはshapeId、textまたはparagraphsのどちらか一方からなるフィールド一覧。paragraphsは各項目にtext、kind=Plain/Bullet/Numbered、level=0〜4、番号開始時だけstart_atを指定します。名前が重複する場合はshapeIdを指定します。")]
        IReadOnlyList<TemplateField> fields,
        CancellationToken cancellationToken,
        [Description("LibreChatの現在のメッセージへ添付したテンプレートPPTXのfile_id。省略時は今回の添付内で最新のPPTX。")]
        string sourceFileId = "latest") =>
        ExecuteJobSubmissionAsync(
            "pptx_populate_template",
            () => jobs.SubmitPopulateTemplateAsync(
                callerContext.GetRequired(),
                sourceFileId,
                fields,
                cancellationToken));

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

        return await ExecuteJobSubmissionAsync(
            "pptx_create_deck",
            () => jobs.SubmitCreateDeckAsync(
                callerContext.GetRequired(),
                sourceFileId,
                slides,
                cancellationToken)).ConfigureAwait(false);
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

        return await ExecuteJobSubmissionAsync(
            "pptx_refine_deck",
            () => jobs.SubmitRefineDeckAsync(
                callerContext.GetRequired(),
                jobId,
                requestedRevisions,
                cancellationToken)).ConfigureAwait(false);
    }

    [McpServerTool(Name = "pptx_start_visual_deck", Destructive = true),
     Description("新規Visual Deckのサーバー側ドラフトを1回だけ開始し、タイトル、ページ数、theme/design、テンプレートを固定します。Design Briefはpptx_validate_design_brief、または必要時のpptx_prepare_design_brief→ユーザー選択→pptx_apply_design_brief_actionで確定し、そのbriefIdを渡してtheme/designを省略します。選択カードpending中や既にstartedのbriefを別資料へ再利用するstartは拒否します。成功済みデッキの見た目はページ単位で修正し、ユーザーが別資料を明示した場合だけuserRequestedNewWorkflow=trueにします。この段階では本文/PPTXを生成しません。")]
    public static async Task<object> StartVisualDeck(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        JobService jobs,
        DesignBriefService designBriefs,
        [Required, Description("資料全体のタイトル。1〜160文字。")]
        string title,
        [Required, Description("完成時の総ページ数。1〜50。後から変更しないため、構成を決めてから指定します。")]
        int expectedSlideCount,
        [Description("任意の全体テーマ。presetはmidnight/aurora/sunset/forest/minimal/ocean/berry/clay/cyber。色roleはprimary/secondary/accent/background/surface/text/mutedText/positive/warning/criticalと1〜8色のdataSeriesColors、font roleはheadingFontFace/bodyFontFaceを使います。legacy fontFaceは両roleのfallbackです。明示値はテンプレート抽出値より優先され、Design BriefのbriefIdを使う場合はtheme自体を省略します。")]
        VisualThemeSpec? theme = null,
        [Description("任意の資料サブジェクト。最大240文字。")]
        string? subject = null,
        [Description("言語コード。省略時ja-JP。")]
        string language = "ja-JP",
        [Description("任意の全体デザイン。style、density、motifを指定します。")]
        VisualDesignSpec? design = null,
        [Description("生成前に固定するテンプレート。既定はdefault、テンプレートなしはnone、添付はlatestまたはfile_id。完成時に変更できません。")]
        string templateSourceFileId = "default",
        [Description("通常はauto。添付テンプレートの特定レイアウトを使う場合だけpptx_analyzeの正確なlayout_id。生成前に固定されます。")]
        string templateLayoutId = "auto",
        [Description("既に成功したデッキとは別の新しい資料を、ユーザーが明示的に依頼した場合だけtrue。見た目の修正や自動リトライではfalseのままにします。")]
        bool userRequestedNewWorkflow = false,
        [Description("任意。同じ利用者・会話でpptx_validate_design_brief、またはカード選択後のpptx_apply_design_brief_actionが返した期限内brief_id。pendingカードがあればRequireDesignBrief=falseでも省略不可です。指定時はtheme/designを省略し、profileのtemplate_sourceとexpectedSlideCountを一致させます。started briefは別資料へ再利用できません。")]
        string? briefId = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var caller = callerContext.GetRequired();
            var designBrief = designBriefs.AuthorizeForStart(
                caller,
                briefId,
                expectedSlideCount,
                templateSourceFileId,
                theme,
                design,
                userRequestedNewWorkflow);
            var permission = await jobs.AuthorizeVisualDeckStartAsync(
                caller,
                userRequestedNewWorkflow,
                cancellationToken).ConfigureAwait(false);
            var reserved = designBriefs.ReserveStart(caller, designBrief);
            try
            {
                var started = drafts.Begin(
                    caller,
                    title,
                    expectedSlideCount,
                    designBrief?.Theme ?? theme,
                    subject,
                    language,
                    designBrief?.Design ?? design,
                    templateSourceFileId,
                    templateLayoutId,
                    permission.AllowSubmittedReplacement,
                    designBrief);
                designBriefs.MarkStartSucceeded(caller, designBrief?.BriefId);
                return started;
            }
            catch
            {
                if (reserved)
                {
                    designBriefs.ReleaseStartReservation(caller, designBrief);
                }

                throw;
            }
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_start_visual_deck", exception);
        }
    }

    [McpServerTool(Name = "pptx_add_visual_slides_to_draft", Destructive = true),
     Description("新規Visual Deckのドラフトへ、順番どおりの完成済みVisualSlideSpecを1〜4ページ追加します。visual-v7-author-htmlでは各ページのHTML/CSSを最後まで完成させ、2〜4ページ、残りが1ページだけならその1ページを送ります。直前の結果の範囲内で、HTML/CSSを途中切れにせず完成できる最大のbatchを優先します。startSlideNumberは省略でき、サーバーが現在の末尾から自動計算します。各ページのspeakerNotesへpurpose（このページで訴えること）とtalkScript（読み上げ原稿）を入れます。ノートはPowerPointの発表者ノートへ保存され、スライド面へ描画されません。Title/Agenda/Section/Bullets/Statement/Cards/Metrics/Comparison/StructuredBrief/Scorecard/DataTable/Media/NativeDiagram/MusicScore/Process/Timeline/Matrix/Funnel/Roadmap/Chart/Dashboard/Quote/Closingを内容に応じて使い分けます。NativeDiagramはdiagram.kind=tree|flow|cycle|concentric|networkと上限付きnodes/edgesを使い、座標・SVG/XMLは使いません。Mediaはvariant=splitと、pptx_register_uploaded_image_assetが返した同じ会話のmedia.assetId、cropIntent、textPositionを必須とし、空欄・仮画像・URL・pathを完成扱いしません。StructuredBriefはstructuredBrief.sections、Scorecardは評価軸×選択肢、DataTableはdataTableの編集可能なcolumns/rows/cells（header/cell textは明示改行なし）、MusicScoreはmusicScoreの五線譜とTABを使います。slide単位のdensityはairy/balanced/detailedでdeck既定を上書きできます。variantはautoのほか、MediaまたはBullets 4件以上かつtakeawayなしのsplit、Metrics正確に3件またはCards 3〜4件のspotlight、StructuredBrief 3 sectionsのeditorial、Processのloop、Timeline/Roadmapのstepped、Funnelのpyramidだけを使えます。prepared visual objectはslide.visualObjectsへ最大3 IDを明示できます。Design Brief利用時は省略すればAsset Planの同一IDをサーバーが補完し、明示した不一致は拒否します。Design Briefを使うdraftでは全slideへ計画済みrecipeIdをコピーし、recipeのkind/density/variantから変えません。6枚以上では4種類以上の構図を計画し、受理済みページを再送せず、remaining_slide_countが0になるまでfinishしません。")]
    public static object AddVisualSlidesToDraft(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        [Required, Description("pptx_start_visual_deckが返したdraft_id。")]
        string draftId,
        [Required, Description("追加する完成済みの連続1〜4ページ。visual-v7-author-htmlでは各ページのHTML/CSSを完結させ、2〜4ページ、残りが1ページだけならその1ページをまとめます。4ページへ届かせるためにHTML/CSSを省略しません。各slideのspeakerNotesへ一行240文字以内のpurposeと1200文字以内のtalkScriptを指定します。Mediaはmedia.assetId、cropIntent=contain|cover|focalCenter|focalLeft|focalRight、textPosition=left|rightを指定し、画像が未登録ならMediaを使いません。NativeDiagramはcycle 3〜6、concentric 2〜4、network 3〜9、tree/flow 3〜12 nodes、edges最大18です。visualObjectsは同slide向けに準備したopaque assetIdを最大3件だけ指定できます。Design Briefで計画済みなら省略時にサーバーが補完します。Metricsはmetricsを2〜6件、Dashboardはmetricsを2〜4件とchart、Cardsはcardsを3〜6件、Comparisonはpanelsを2〜3件、StructuredBriefはsectionsを2〜3件・合計900文字以内、Scorecardはoptionsを2〜4件、criteriaを2〜6行かつ各cells数をoptions数と一致させます。DataTableはdataTable.columnsとdataTable.rowsを使い、全rowのcells数をcolumns数と一致させ、header/cell textを明示改行なしの1行にします。airyは最大4列×6行、balancedは5列×8行、detailedは6列×10行です。MusicScoreはmusicScoreへ1〜8小節・最大64イベントを指定します。Process/Timeline/Funnel/Roadmapはstepsを3〜6件、Matrixはquadrantsを正確に4件です。任意のdensityと、Design Brief利用時は必須のrecipeIdをslide直下に指定します。")]
        IReadOnlyList<VisualSlideSpec> slides,
        [Description("省略推奨。明示する場合はこのバッチの先頭ページ番号で、直前のnext_slide_numberと一致させます。")]
        int? startSlideNumber = null)
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
     Description("全ページを追加済みのVisual DeckドラフトをPPTXとして生成します。remaining_slide_count=0になった後にdraftId付きで1回だけ呼びます。テンプレート、theme、designはpptx_start_visual_deckで固定済みなので、この段階では変更しません。生成成功後はこのツールやpptx_start_visual_deckを再実行せず、全ページを視覚確認して問題ページだけをpptx_refine_visual_slideで直してください。")]
    public static Task<object> FinishVisualDeckAsync(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        JobService jobs,
        [Required, Description("全ページ追加後のVisual DeckドラフトID。")]
        string draftId,
        [Description("互換用。通常は省略します。指定した場合は開始時に固定したtemplateSourceFileId（true=default、false=none）と一致しなければエラーになります。")]
        bool? useDefaultTemplate = null,
        CancellationToken cancellationToken = default) =>
        FinishVisualDraftAsync(
            "pptx_finish_visual_deck",
            callerContext,
            drafts,
            jobs,
            draftId,
            useDefaultTemplate.HasValue ? (useDefaultTemplate.Value ? "default" : "none") : null,
            null,
            cancellationToken);

    [McpServerTool(Name = "pptx_finish_branded_visual_deck", Destructive = true),
     Description("全ページを追加済みのVisual Deckドラフトを、既定またはアップロード済み企業テンプレートのマスター、ロゴ、フッター、ページ設定へ合成します。通常の新規資料はpptx_finish_visual_deckが既定テンプレートを自動適用するため、主に添付した別テンプレートを明示的に使う場合に呼びます。事前に添付を解析し、remaining_slide_count=0になったドラフトIDを指定してください。templateLayoutIdは通常autoです。")]
    public static Task<object> FinishBrandedVisualDeckAsync(
        CallerContextAccessor callerContext,
        VisualDeckDraftService drafts,
        JobService jobs,
        [Required, Description("全ページ追加後のVisual DeckドラフトID。")]
        string draftId,
        [Description("通常は省略し、開始時に固定したautoまたは正確なtemplateLayoutIdを使います。指定する場合は開始時の値と完全一致させます。")]
        string? templateLayoutId = null,
        [Description("通常は省略。指定する場合は開始時に固定したtemplateSourceFileIdと一致させます。")]
        string? sourceFileId = null,
        CancellationToken cancellationToken = default) =>
        FinishVisualDraftAsync(
            "pptx_finish_branded_visual_deck",
            callerContext,
            drafts,
            jobs,
            draftId,
            sourceFileId,
            templateLayoutId,
            cancellationToken);

    [McpServerTool(Name = "pptx_insert_visual_slides", Destructive = true),
     Description("成功したVisual DeckまたはブランドVisual Deckへ新しいページを追加します。ユーザーが追加・挿入・末尾へ追加を依頼した場合は、新しいVisual Deckドラフトで資料全体を作り直さず、このツールへ追加分の完全なVisualSlideSpecだけを渡してください。新規ページにはspeakerNotesのpurposeとtalkScriptも含めます。jobIdは通常latest、afterSlideNumberは省略時に末尾追加、0なら先頭追加、正の値ならその1始まりページの直後へ挿入します。既存の資料タイトル、テーマ、design、全ページ、企業テンプレートと選択レイアウトはサーバー側で継承します。通常はサーバー内で完了まで待って最終statusと成果物を返すため、Succeededなら状態確認ツールを呼ばず全ページを視覚確認してください。30秒以内に完了せずqueuedを返した場合だけjob_idをpptx_wait_for_jobで待ちます。")]
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
     Description("互換用のVisual Deck差分修正ツールです。サーバーは1回につき問題ページ1枚だけを受理し、最大3巡を強制します。通常はjobId=latestを自動解決して30秒待機するpptx_refine_visual_slideを優先してください。ブランドVisual Deckでは元の企業テンプレートと選択レイアウトを維持します。Design Brief / Brand Profile-bound deckの差し替えslideは、元ページのrecipeId、kind、実効density、variantを完全に保持します。prepared visual objectのvisualObjectsとspeakerNotesは省略すればサーバーが元ページの値を継承します。ノートを直す場合だけ新しいpurposeとtalkScriptを明示します。visualObjectsの異なるIDへの変更は拒否します。資料全体を再送してはいけません.")]
    public static async Task<object> RefineVisualDeckAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        CancellationToken cancellationToken,
        [Required, Description("動作上必須。直前の成功したpptx_finish_visual_deckまたはpptx_finish_branded_visual_deckが返したjob_id。")]
        string jobId = "",
        [Required, Description("動作上必須。要素数を正確に1件とし、1始まりのslide_numberと差し替え後の完全なslideを含めます。")]
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
     Description("成功したVisual Deck／ブランドVisual Deck、またはvisual_authored_html_invalid等の回復可能なHTML/CSS検証エラーで失敗したVisual Deckの既存ページを1枚だけ差し替えます。失敗jobの復旧ではエラーに示されたページだけを完全なHTML/CSSへ直し、元のjobIdを明示します。新しいDesign Briefやdraftを開始しません。ページ追加には使わずpptx_insert_visual_slidesを使ってください。Bedrock/Claudeでの自動視覚リフレクションではこのツールを優先します。revisionにslide_numberと差し替え後の完全なslideを必ず含めます。Design Brief / Brand Profile-bound deckでは、元ページのrecipeId、kind、実効density、variantを完全に保持します。prepared visual objectのvisualObjectsとspeakerNotesは省略すればサーバーが元ページの値を継承します。ノートを直す場合だけ新しいpurposeとtalkScriptを明示し、visualObjectsの異なるIDへの変更は拒否します。別recipeや別構図へ変更しません。成功済み資料はjobId=latestを使えますが、失敗jobの復旧では返された正確なjobIdを使います。通常はサーバー内で完了まで待って最終statusと成果物を返すため、Succeededなら状態確認ツールを呼ばず次のページへ進んでください。30秒以内に完了せずqueuedを返した場合だけjob_idをpptx_wait_for_jobで待ちます。複数ページを1ページずつ直すと修正が累積します。同じページの再修正で次巡へ進み、サーバーが全体を最大3巡に制限します。上限後は全体を再作成しません.")]
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
     Description("PowerPointジョブの現在状態を待たずに1回だけ取得します。障害復旧や即時確認用です。待機中・実行中ジョブの通常フローでは、短間隔でこのツールを反復せずpptx_wait_for_jobを使ってください。完了済み解析の巨大な結果本文は二重返却せず、pptx_wait_for_jobで取得します。jobId=latestでは同じ利用者・会話の直近ジョブを安全に選びます。")]
    public static async Task<object> GetJobAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("PowerPoint MCPが返したjob_id、または同じ会話の直近ジョブを選ぶlatest。")]
        string jobId,
        CancellationToken cancellationToken)
    {
        try
        {
            var job = await jobs.GetAsync(callerContext.GetRequired(), jobId, cancellationToken).ConfigureAwait(false);
            return PrepareGetJobResult(job);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_get_job", exception);
        }
    }

    internal static object PrepareGetJobResult(JobView job)
    {
        if (JobService.IsRecoverableAuthoredVisualFailure(job.Status, job.ErrorCode, job.ErrorMessage))
        {
            return PrepareRecoverableVisualFailure(job);
        }
        if (job.Kind != JobKind.Analyze || job.Status != JobState.Succeeded)
        {
            return job;
        }

        return new
        {
            job_id = job.JobId,
            kind = job.Kind,
            status = job.Status,
            progress_percent = job.ProgressPercent,
            created_at = job.CreatedAt,
            completed_at = job.CompletedAt,
            result_omitted = true,
            instruction =
                "The completed analysis body is omitted from pptx_get_job to avoid duplicating a large result. " +
                "If pptx_wait_for_job already returned this job_id, do not call either status tool again; " +
                "use that analysis and call pptx_replace_text now. Only if this turn has never received the analysis body, " +
                "call pptx_wait_for_job once with this job_id to retrieve it.",
            artifacts = job.Artifacts,
            error_code = job.ErrorCode,
            error_message = job.ErrorMessage,
            visual_root_job_id = job.VisualRootJobId,
            visual_revision_round = job.VisualRevisionRound,
            visual_revised_slides_in_round = job.VisualRevisedSlidesInRound,
        };
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
            var job = await jobs.WaitAsync(
                callerContext.GetRequired(),
                jobId,
                TimeSpan.FromSeconds(waitSeconds),
                cancellationToken).ConfigureAwait(false);
            return PrepareWaitForJobResult(job);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError("pptx_wait_for_job", exception);
        }
    }

    internal static object PrepareWaitForJobResult(JobView job)
    {
        if (JobService.IsRecoverableAuthoredVisualFailure(job.Status, job.ErrorCode, job.ErrorMessage))
        {
            return PrepareRecoverableVisualFailure(job);
        }
        if (job.Kind != JobKind.Analyze || job.Status != JobState.Succeeded)
        {
            return job;
        }

        if (!DeliveredAnalysisResults.TryAdd(job.JobId, 0))
        {
            return PrepareGetJobResult(job);
        }

        return job.Result is { } result ? result : PrepareGetJobResult(job);
    }

    private static object PrepareRecoverableVisualFailure(JobView job) => new
    {
        job_id = job.JobId,
        kind = job.Kind,
        status = job.Status,
        progress_percent = job.ProgressPercent,
        created_at = job.CreatedAt,
        completed_at = job.CompletedAt,
        artifacts = job.Artifacts,
        error_code = job.ErrorCode,
        error_message = job.ErrorMessage,
        visual_root_job_id = job.VisualRootJobId,
        visual_revision_round = job.VisualRevisionRound,
        visual_revised_slides_in_round = job.VisualRevisedSlidesInRound,
        instruction =
            "This failed job is recoverable without rebuilding the deck. Call pptx_refine_visual_slide once with this exact job_id and replace only the slide identified in error_message with corrected complete HTML/CSS. Do not validate another Design Brief, call pptx_start_visual_deck, or resend unaffected slides.",
    };

    [McpServerTool(Name = "pptx_get_preview_images", ReadOnly = true, Idempotent = true),
     Description("成功済みジョブのスライド画像をClaude自身の視覚確認用に返します。全ページ確認はスライド番号順の連続4枚ずつ、最後の端数だけ1〜3枚で取得し、同じjobの同じページを重複取得しません。文字切れ・重なり・可読性に加え、見出しだけで話が追えるか、読む順序が一意か、1ブロック1論点か、強調色が概ね15%以内か、本文が9pt未満に見えないか、単調な構図、余白、整列、コントラスト、密度、バランス、全体一貫性を確認してください。厳密なプレースホルダー資料はpptx_refine_deck、白紙またはブランドVisual Deckはpptx_refine_visual_slideへ問題ページを1枚ずつ渡します。最大3巡で収束させ、このツールを呼ばずに視覚確認済みと述べてはいけません.")]
    public static async Task<CallToolResult> GetPreviewImagesAsync(
        CallerContextAccessor callerContext,
        JobService jobs,
        [Description("PowerPoint MCPが返した成功済みjob_id。")]
        string jobId,
        [Description("今回確認する1始まりのスライド番号。重複なしで1〜4件。全ページを複数回に分けて指定します。")]
        IReadOnlyList<int> slideNumbers,
        CancellationToken cancellationToken)
    {
        try
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
                Text = "Evaluate the returned slides for clipping, overflow, overlap, legibility, spacing, alignment, contrast, hierarchy, density, balance, visual variety, and consistency. Also verify that the heading sequence tells the story without body text, reading order is unambiguous, each block has one main point, emphasis color occupies roughly 15% or less, and content text does not appear below 9 pt. For a placeholder template deck use pptx_refine_deck; for a visual or branded visual deck use pptx_refine_visual_slide with one complete replacement slide at a time. Never resend, restart, or rebuild the complete deck after a successful generation. The server permits at most three visual refinement rounds. If clipping survives one revision, shorten visible copy or simplify the structure instead of merely tightening spacing or reducing font size.",
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
        catch (PptxValidationException exception)
        {
            var error = CreateValidationError("pptx_get_preview_images", exception);
            return new CallToolResult
            {
                Content = [new TextContentBlock { Text = JsonSerializer.Serialize(error) }],
                StructuredContent = JsonSerializer.SerializeToElement(error),
            };
        }
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
        string? requestedTemplateSourceFileId,
        string? requestedTemplateLayoutId,
        CancellationToken cancellationToken)
    {
        var caller = callerContext.GetRequired();
        var acquired = false;
        try
        {
            var submission = drafts.AcquireForSubmission(
                caller,
                draftId,
                requestedTemplateSourceFileId,
                requestedTemplateLayoutId);
            if (submission.ExistingJobId is not null)
            {
                return await jobs.GetAsync(caller, submission.ExistingJobId, cancellationToken).ConfigureAwait(false);
            }

            acquired = true;
            var receipt = await SubmitLockedVisualDeckAsync(
                jobs,
                caller,
                submission,
                cancellationToken).ConfigureAwait(false);
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

    private static Task<JobReceipt> SubmitLockedVisualDeckAsync(
        JobService jobs,
        CallerContext caller,
        VisualDeckDraftSubmission submission,
        CancellationToken cancellationToken)
    {
        if (string.Equals(submission.TemplateSourceFileId, "none", StringComparison.OrdinalIgnoreCase))
        {
            return jobs.SubmitVisualDeckAsync(caller, submission.Deck!, false, cancellationToken);
        }

        if (string.Equals(submission.TemplateSourceFileId, "default", StringComparison.OrdinalIgnoreCase)
            && string.Equals(submission.TemplateLayoutId, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return jobs.SubmitVisualDeckAsync(caller, submission.Deck!, true, cancellationToken);
        }

        return jobs.SubmitBrandedVisualDeckAsync(
            caller,
            submission.TemplateSourceFileId,
            submission.Deck!,
            submission.TemplateLayoutId,
            cancellationToken);
    }

    internal static async Task<object> ExecuteJobSubmissionAsync(
        string toolName,
        Func<Task<JobReceipt>> submit)
    {
        try
        {
            return await submit().ConfigureAwait(false);
        }
        catch (PptxValidationException exception)
        {
            return CreateValidationError(toolName, exception);
        }
    }

    private static ToolValidationError CreateValidationError(string tool, PptxValidationException exception)
    {
        var instruction = exception.Code switch
        {
            "visual_deck_already_completed" =>
                "Do not call pptx_start_visual_deck again. Use pptx_get_preview_images, pptx_refine_visual_slide, or pptx_insert_visual_slides. Set userRequestedNewWorkflow=true only after an explicit user request for a separate deck.",
            "visual_deck_generation_in_progress" or "visual_deck_operation_in_progress" =>
                "Wait for the existing job with pptx_wait_for_job. Do not start or branch another full-deck job.",
            "visual_deck_recovery_limit_reached" or "visual_refinement_limit_reached" =>
                "Stop calling PowerPoint mutation tools and report the latest successful result or failure to the user.",
            "visual_deck_failed_page_refinement_required" =>
                "Do not validate another Design Brief or call pptx_start_visual_deck. Use the exact failed job ID in message with pptx_refine_visual_slide and replace only the page identified by its HTML/CSS validation error.",
            "visual_deck_job_superseded" =>
                "Call the same page-level operation once with jobId=latest so all accepted changes are preserved.",
            "visual_creative_direction_locked" =>
                "Finish the existing draft with its locked template, theme, and design. Do not restart the whole deck to change appearance.",
            "visual_draft_not_found" or "visual_draft_expired" =>
                "Stop this turn. Do not call pptx_add_visual_slides_to_draft or a finish tool with this draftId again. Tell the user that the draft is unavailable or expired; a fresh Design Brief and pptx_start_visual_deck are required to generate the deck again.",
            "visual_draft_not_editable" or "visual_draft_already_submitted" =>
                "Do not retry this draft mutation. Use the existing submitted job when available, or report that the draft can no longer be edited.",
            "visual_draft_already_active" =>
                "Do not validate another Design Brief and do not call pptx_start_visual_deck again. Continue the exact active draft_id and next slide stated in message with pptx_add_visual_slides_to_draft.",
            "visual_deck_job_not_refinable" =>
                "Do not retry refinement or restart the deck. Report the exact job error, unless a separate successful Visual Deck job is already available for page-level refinement.",
            "design_brief_required" =>
                "Call pptx_get_design_catalog, resolve only material user questions, call pptx_validate_design_brief, and then retry pptx_start_visual_deck once with the returned briefId.",
            "design_catalog_profile_required" =>
                "Call pptx_get_design_catalog once without filters, choose one exact profileId and style_directions entry, then call it exactly once more with that profileId and styleDirectionId. A mixed-purpose deck must omit purpose and density.",
            "design_brief_expired"
                or "design_brief_not_found"
                or "design_brief_identifier_invalid"
                or "brand_profile_version_mismatch" =>
                "Refresh pptx_get_design_catalog and validate a new Design Brief before starting. Do not reuse the expired or stale briefId.",
            "design_brief_action_expired"
                or "design_brief_action_not_found"
                or "design_brief_action_state_invalid" =>
                "Refresh pptx_get_design_catalog and prepare a new Design Brief card. The old opaque card identifiers are no longer usable.",
            "design_brief_action_tampered" or "design_brief_action_identifier_invalid" =>
                "Ignore the untrusted action. Accept only a server-issued option from the current card; replace the card only after an explicit user request.",
            "design_brief_action_replayed"
                or "design_brief_action_already_started" =>
                "Do not apply another option. Continue the already selected brief or the existing Visual Deck draft.",
            "design_brief_choice_pending" or "design_brief_not_confirmed" =>
                "Stop this turn and wait for the user action from the existing Design Brief card. Do not validate, prepare, or start another brief.",
            "design_brief_choice_not_pending" =>
                "Do not retry cancellation. This conversation has no unselected card; continue the current validated or selected workflow, or prepare a new card only if the user requests one.",
            "design_brief_choice_cancel_forbidden" =>
                "Do not retry cancellation or choose another option. Continue only the already selected brief or existing Visual Deck draft.",
            "design_brief_selection_required"
                or "design_brief_selection_superseded"
                or "design_brief_choice_already_applied" =>
                "Use only the briefId already returned by pptx_apply_design_brief_action. Do not bypass or replace the applied selection before start.",
            "design_brief_card_choice_required"
                or "design_brief_alternative_has_no_visual_difference"
                or "design_brief_no_photo_has_no_visual_difference" =>
                "Do not show a one-choice or cosmetic-only card. Use pptx_validate_design_brief for the safe recommendation, or supply a genuinely different executable recipe/style choice.",
            "design_brief_start_in_progress"
                or "design_brief_action_start_in_progress" =>
                "Wait for the current start call. Do not prepare, validate, apply, or start another Design Brief concurrently.",
            "design_brief_start_state_invalid" or "design_brief_start_state_changed" =>
                "Do not retry start with stale state. Re-apply the current valid card option, or prepare a new card after the old choice expires or is explicitly replaced.",
            "design_brief_capacity_reached" or "design_brief_user_capacity_reached" =>
                "Stop retrying. Wait for an existing brief/card to complete or expire, then continue once.",
            "design_brief_ui_resource_too_large" =>
                "Do not retry the same card. Use direct pptx_validate_design_brief for the safe recommendation and report that the comparison card could not be rendered.",
            "design_brief_already_started" =>
                "For an explicitly requested separate deck, validate or select a new Design Brief. Keep flag=false only for an idempotent retry of the existing active draft.",
            "design_brief_confirmation_required" or "design_brief_questions_unresolved" =>
                "Ask at most the material unresolved questions or select a safe explicit fallback, then validate the finalized brief with no unresolved questions.",
            "asset_plan_omission_invalid" =>
                "For that item use exactly preferred_medium=none, acquisition=none, fallback=none, status=omitted, and license_status=notRequired; omit approved_asset_collection_id, attribution_ref, crop_intent, aspect_ratio, and text_safe_area; select a recipe whose required_asset_roles is empty. Never pair acquisition=none with noAssetLayout.",
            "visual_slide_recipe_mismatch"
                or "visual_slide_recipe_kind_mismatch"
                or "visual_slide_recipe_density_mismatch"
                or "visual_slide_recipe_variant_mismatch" =>
                "Copy the exact planned recipeId and match its semantic kind, effective density, and implemented variant. Do not start a new draft.",
            "brand_profile_insert_requires_design_brief" =>
                "Do not insert slides into this Brand Profile-bound deck in phase 1. Explain that inserted pages need a newly validated Asset Plan and recipe contract; keep the latest successful deck unchanged.",
            "source_preview_before_text_edit_forbidden" =>
                "Use the complete structured text already returned by pptx_analyze and start pptx_replace_text now. Do not call pptx_render_preview, pptx_get_preview_images, pptx_get_job, pptx_wait_for_job, or pptx_analyze again for the unmodified source. This guard has no model-controlled override. If the user explicitly asks to inspect the original presentation in a later message, call pptx_render_preview directly in that later message before any new analysis.",
            "text_edit_workflow_already_started" or "text_edit_job_superseded" =>
                "Do not restart pptx_replace_text from the source presentation. Use the exact latest job_id and final_batch_submitted state in message. If more replacements are needed, pass that job_id as previousJobId; if the final batch is already submitted and no correction remains, use that final job's artifacts and finish the user response.",
            "preview_slide_invalid" or "preview_selection_invalid" =>
                "Use only the valid one-based slide numbers stated in message. Continue with the remaining valid slides and never guess slide numbers beyond this job's slide count. Do not call pptx_get_job or restart the workflow.",
            "file_not_found" =>
                "Tell the user that no PPTX attached to the current request is available and ask them to attach it once. State that the supported PowerPoint upload format is .pptx only; do not suggest the legacy .ppt format. Do not call another PowerPoint job tool until an upload succeeds.",
            "invalid_file_id" =>
                $"Use latest or the exact opaque file_id supplied by LibreChat when calling {tool}. Never use a filename, URL, or local path.",
            "ambiguous_file_id" =>
                $"Call {tool} once with the exact opaque file_id returned by LibreChat. Do not use a filename or path.",
            "file_size_out_of_range" =>
                "Stop and tell the user in the user's language that the PPTX is empty or exceeds the accepted upload size. Ask them to attach a non-empty PPTX within the size limit shown in message. Do not retry the unchanged file.",
            "slide_count_out_of_range" =>
                "Stop and tell the user in the user's language that the PPTX must contain a supported number of slides as shown in message. Ask them to attach a corrected PPTX. Do not retry the unchanged file.",
            "external_relationship" =>
                "Stop and reply in the user's language with a simple explanation of what was blocked. Explain that an external image means the slide points to an image stored on a company server or website instead of containing the image itself, so it disappears if the source becomes unavailable. Explain that OLE is the PowerPoint feature for handling an Excel table, chart, or Word document, and a linked OLE object may open or read the original file. Explain that a shared-file reference points to another file on a shared folder, network drive, or SharePoint and reflects that file's contents. State that ordinary HTTP or HTTPS web hyperlinks are allowed and are not treated as these external resources; a non-web hyperlink such as a local or network file link is blocked. Ask the user to remove or embed the blocked reference and upload the corrected PPTX before retrying. Do not retry the unchanged file.",
            "active_content" =>
                "Stop and explain in the user's language that the PPTX contains a macro or ActiveX component, which is a program-like feature that can run actions and is not accepted for safety. Ask the user to save and attach a normal .pptx with that active content removed. Do not retry the unchanged file.",
            "invalid_pptx"
                or "invalid_zip"
                or "invalid_xml"
                or "zip_entry_limit"
                or "zip_expansion_limit"
                or "zip_compression_ratio"
                or "zip_path_traversal" =>
                "Stop and explain in the user's language that the file is damaged, is not a readable PPTX, or exceeds a package safety limit, using the validation code and message to identify which. Ask the user to reopen and save it as a normal .pptx or attach a corrected copy. Do not retry or modify the rejected file on the user's behalf.",
            _ => $"Correct the field named in message and call {tool} again. Do not repeat the same invalid input.",
        };
        return new(
            "invalid_input",
            tool,
            exception.Code,
            exception.Message,
            instruction);
    }
}
