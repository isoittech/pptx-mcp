# AGENTS.md

## プロジェクト概要

- LibreChat から利用する PowerPoint 生成・更新 MCP サーバーである。
- .NET 8、公式 MCP C# SDK、Open XML SDK を使用する。
- PPTX のバイナリや全プレビュー画像を MCP 応答へ埋め込まず、非同期ジョブと署名付き成果物 URL を使う。

## よく使うコマンド

- `docker compose build`: アプリとテスト用イメージをビルドする。
- `docker compose run --rm test`: 全テストを実行する。
- `cd visual-renderer && npm audit --omit=dev`: 白紙生成レンダラーの依存脆弱性を確認する。
- `docker compose up -d pptx-mcp`: MCP サーバーを起動する。
- 提供テンプレートや顧客資料は `tmp/` に置き、Gitへ追加しない。

## アーキテクチャ方針

- MCP ツールは受付・状態取得を担当し、重い処理は最大3並列のバックグラウンドジョブで行う。
- ユーザーと会話の識別子は LibreChat が付与する信頼済み HTTP ヘッダーから取得し、ツール引数から受け取らない。
- 外部からローカルパスを受け取らず、不透明な `file_id`、`job_id`、`artifact_id` のみを扱う。
- PPTX 操作は `IPresentationEngine` の背後に置き、Open XML 実装と将来の商用エンジンを交換可能にする。
- LibreOffice は表示確認用レンダリングに限定し、編集には使用しない。
- 新規生成の主経路は`visual-v7-author-html`とし、会話モデルが各ページの完成した静的HTML/CSSを設計する。サーバーは構図を固定レイアウトへ置換せず、allowlist検証、CSS scope付与、承認済み素材解決、Chromium描画、`dom-to-pptx`変換、企業テンプレート合成だけを担当する。意味layout、recipe、density、variantは計画・監査用であり、見た目の正本は`authoredHtml`とする。visible contentを旧renderer用の意味payloadへ二重記載させず、保存済み再試行に残っていてもrenderer固有のgeometry検査へ使わない。Office系の文字メトリクス差に備え、独立領域間に20px以上、文字量の多い領域に概ね15%の高さ予備を設計させる。詳細はADR 0022を参照する。
- `visual-v7-author-html`は全ページをDOM変換し、PptxGenJS互換経路へfallbackしない。`visual-v6-dom`以前は保存済みlineageの互換再生成に限定し、新規品質の基準にしない。
- モデル生成HTML/CSSは静的な限定言語として扱う。JavaScript、event属性、外部URL、data URI、ローカルpath、任意SVG/XML、iframe/object/embed/form、`@` rule、`url()`、fixed配置、疑似要素content、animation、transition、shadowを拒否する。HTMLはparse5のtag/attribute allowlistで再構築し、CSSは全selectorが`.slide`から始まる場合だけページ固有scopeへ変換する。
- Lucideはモデルが`data-pptx-icon="承認済み意味ID"`を指定し、サーバー側allowlistから`react-icons/lu`へ解決する。モデルにReact code、SVG本文、package名を生成させない。画像は事前登録済みopaque IDを`data-pptx-asset`で参照し、宣言IDとscope、SHA-256、PNGをサーバーが検証する。
- `VisualDeckSpec`の意味レイアウト、recipe、theme、design、variantは、内容構成と品質検査に使う部品語彙であり金型ではない。会話モデルは内容に合うHTML構造、CSS配置、図形、カード、表、余白を自ら設計する。CoverageMap、TransformationEvidence、ArtifactShowcase、GanttScheduleの選択基準は`docs/visual-component-catalog.md`を参照する。
- DOMの文字・図形はマットにし、`box-shadow`と`text-shadow`を使わない。本文は原則14pt以上とし、出典・URL・発行日は12ptまで許容する。dom-to-pptxでは24pxが約14.4ptになるため、モデル指定の`font-size`はpx単位かつ24px以上を検証し、出典要素自身の`data-pptx-role="source-meta"`だけ20pxまで許容する。既定企業テンプレート本文は、semanticなtitle／subtitleと同じ表示文字列を`body-title`見出しと単一項目`body-claim`実箇条書きへ分離し、モデル自身が最終値の50px／27pxで構図を設計する。serverの最終CSSは30pt相当／16pt相当とAccent 2を保証し、両者を同じヘッダー群として8px間隔に正規化するが、小さい文字で設計して後から拡大する使い方を許可しない。`body-claim`の左paddingは0へ正規化し、dom-to-pptxがPowerPointの上insetへ誤写像して見た目の空白を増やすことを防ぐ。主張と次の独立ブロックの間は20px以上空ける。箇条書きは`ul`/`li`からDrawingMLの実箇条書きへ変換し、装飾用の点と通常テキストで代用しない。既定テンプレートのフッター上に安全帯を残し、過密時は縮小せず、文章整理またはページ分割を選ぶ。
- 既定企業テンプレートの表紙は、sample coverの背景を活かした透明overlayとし、`cover-title` 54pxのタイトルと`cover-subtitle` 27pxの短いサブタイトルだけを表示する。白パネル、機密区分ラベル、対象、論点一覧、長い出典、アイコン、追加画像は本文またはノートへ移す。利用者指定テンプレートにはこの表紙規則を流用しない。
- 楽譜は`MusicScore`へ音高・音価・ウクレレ弦・フレット・指番号を渡し、PowerPointネイティブの線・図形・テキストで五線譜とTABを描く。任意座標を公開せず、音高と調弦・弦・フレットの一致を検証する。音楽記号は同梱したBravura 1.392の輪郭を`custGeom`へ変換し、手描き近似、画像貼付、閲覧側フォント依存へ戻さない。OTFとSIL OFL原文は必ず一緒に更新する。
- 文字量の多い説明は`StructuredBrief`の2〜3セクションへ分け、評価軸×選択肢は`Scorecard`の編集可能なPowerPoint表にする。`density=detailed`はフォント縮小だけで実装せず、外周余白、見出し領域、間隔、罫線、影を一体で切り替える。
- `tone` は意味語の別名またはRGB色を許容し、自然な表現を固定4語へ押し込めない。組み込みアイコンもモデルが実際に使う業務語彙をE2Eで確認して拡張し、未知値を黙って別アイコンへ置換しない。
- テンプレートのリストは `DeckField.paragraphs` / `TemplateField.paragraphs` を使い、1項目1段落で `Plain`、`Bullet`、`Numbered` と0〜4の `level` を指定する。本文へ `■`、`・`、`1.` 等を手入力しない。
- `pptx_analyze` の `theme` はaccent1〜3、light1、dark1、日本語優先の見出し・本文フォントを返す。白紙生成へ企業スタイルを移す場合はこの値を `VisualDeckSpec.theme` に使う。
- 新規Visual Deckは`pptx_start_visual_deck`、最大4ページずつの`pptx_add_visual_slides_to_draft`、finishの順で作る。`visual-v7-author-html`は各ページのHTML/CSSを完結させた連続2〜4ページを受理し、残り1ページだけは単独で受理する。完成HTML/CSSは出力量が大きいため通常は2ページを1batchとし、3〜4ページは短いページだけを完全なtool引数が応答枠へ十分収まる場合に選ぶ。ページ数を増やすためにHTML/CSSを省略しない。最大50ページの完全仕様を一度に要求する公開ツールへ戻さない。`startSlideNumber`は省略可能とし、サーバーが受理済み末尾から追番する。明示値は順序検証にだけ使う。
- start時にテンプレート、テーマ、デザインを固定し、finishでの変更を拒否する。テンプレート抽出値は未指定テーマ項目だけを補完し、明示色・フォントを上書きしない。`pptx_create_deck` は既存プレースホルダー配置への厳密な流し込みが明示された場合だけ使う。
- 導入環境の既定テンプレートは外部マウントと`DefaultTemplateId`で指定し、実PPTX・会社名・ロゴ・固有文言をOSSへ含めない。起動時に検証・解析し、startの`templateSourceFileId=default`で自動適用する。cover/body見本スライド番号を設定した環境では1枚目と2枚目以降へ各layoutをサーバー側で適用する。既定表紙が濃色なら`DefaultTemplateCoverUsesLightForeground`も明示し、1枚目のTitle/Sectionだけ白文字・装飾番号なしで重ねる。添付テンプレートはstartで`latest`または`file_id`、テンプレートなしは`none`を選び、ユーザー指定テンプレートへ既定cover/body設定を流用しない。
- `DefaultTemplateBodyUsesAccent2Headings`を有効にする導入環境では、既定テンプレートの本文タイトルを30pt、主張を別オブジェクトの16pt実箇条書きにし、両方へAccent 2を使う。タイトルへ主張を連結しない。ユーザー指定テンプレートとテンプレートなしには流用しない。
- `UseModelAuthoredHtmlRenderer`は新規draftを`visual-v7-author-html`へ切り替える導入設定で、OSS既定はoffとする。検収環境では`RequireDomOnlyRenderer`も有効にし、全ページの`renderer_usage_by_slide`が`dom-to-pptx`、`fallback_rendered_slide_count`が0であることを確認する。
- Brand Profileは`BrandProfilesRoot/<profile-id>/brand-profile.json`から読み取り専用で起動時に検証し、会社固有値、URL、パスをOSSコードへ含めない。`pptx_get_design_catalog`はsummary＋compactなstyle directionsと、選択profile＋directionの最終detailを各1回だけ返す。`pptx_validate_design_brief`はuser・conversation・profile version/hashへ束縛した期限付き`brief_id`を担当する。`RequireDesignBrief`はOSS既定offとし、導入環境で有効化した場合だけstart前に必須化する。詳細はADR 0013を参照する。
- デザイン方向が結果へ大きく影響し、実効renderer fingerprintが異なる案を2件以上出せる場合だけ、`pptx_prepare_design_brief`→UI Resourceでturn終了→固定`pptx.designBrief.select` intent→`pptx_apply_design_brief_action`の順にする。候補は最大3件、clientからは`choiceSessionId`と`optionId`だけを受け、pending中はoptional構成でもvalidate/startを拒否する。明確な依頼や1案だけなら従来のvalidateへ直行する。カード非表示または利用者がsafe defaultを明示した場合だけ、引数なしの`pptx_cancel_design_brief_selection`で未選択pendingを破棄する。選択済みは取消不可とする。詳細はADR 0015を参照する。
- Brand sample thumbnailは`sample-thumbnails/<sample-id>.png`の検査済みnon-interlaced PNGだけをUI確認用に任意読込する。model/renderer入力やPPTX素材へ使わず、任意URL・JPEG・path・symlink・metadata chunkを拒否する。profile単位ACLはないため、対象deploymentの全利用者へ配布承認済み・非機密・権利確認済みのderivativeだけを登録する。技術検査を権利承認と呼ばない。
- 複数のBrand sample thumbnailは16:9なら960x540程度を基準にする。1profile全体の6Mpx上限は各画像の上限とは別に適用され、1280x720を9枚置く構成は起動時検証で拒否される。
- Design Briefを使うdraftは全ページのAsset Planとlayout recipeを固定し、add時に`recipeId`、意味layout、密度、実装済みvariantを照合する。`userUpload`画像を使う場合は先に`pptx_register_uploaded_image_asset`で会話scope付きassetへ登録し、`ready`、`userProvided`、`fallback=none`、`asset_id`、crop、text safe area、画像必須`Media/split` recipeを固定する。未登録uploadと`approvedLibrary`はnativeDrawまたは画像なしlayoutのfallbackへ切り替える。素材を使わない項目は`preferred_medium=none`、`acquisition=none`、`fallback=none`、`status=omitted`、`license_status=notRequired`の組合せにし、`noAssetLayout`は`userUpload`または`approvedLibrary`の`fallbackSelected`にだけ使う。
- 完成したprofile-bound deckには秘密を含まないDesign Brief監査snapshotとページrecipe契約を保存する。refineでは元ページの`recipeId`、kind、実効density、variantと監査snapshotを保持する。prepared visual object参照は置換slideで省略された場合にjob snapshotからmaterializeし、明示された異なるIDは拒否する。第1段階のinsertは新規Asset Planを検証できないため拒否し、期限付き`brief_id`はjob payloadへ保存しない。
- Design Brief監査には`selection_source=agentDefault|userCard`だけを保存し、choice session、option、nonce、brief IDをjob payloadへ残さない。旧payloadでfieldが欠落する場合は`agentDefault`として読む。
- 生成・編集後は `pptx_get_preview_images` で全ページをClaudeへ渡し、問題時は宣言型仕様を最大3回まで修正する。通常の欠け・重なりに加え、見出しだけで話が追えるか、読む順序が一意か、1ブロック1論点か、強調色が概ね15%以内か、本文が9pt未満に見えないかを確認する。
- Bedrockから白紙資料とブランドVisual Deckを視覚修正する場合は、`pptx_refine_visual_slide` へ完全な差し替えページを1枚ずつ渡し、各成功後に `jobId=latest` で次ページを直す。ジョブへルート・親・修正巡を保存し、古い版からの分岐、一括ページ修正、4巡目をサーバー側で拒否する。同じ欠けが1巡後も残る場合は余白だけを詰めず、表示文の短縮、ノートへの詳細移動、HTML構造の単純化を行う。成功後のstart再実行も拒否し、初回生成失敗時の全体再試行だけ1回許可する。
- 成功済みVisual Deckへページを増やす場合は`pptx_insert_visual_slides`へ追加ページだけを渡す。新規ドラフトを開始せず、既存ページも再送しない。サーバー側で元仕様へ挿入し、ブランド資料では元テンプレートとレイアウトを継承する。第1段階では完成版全体を再レンダリングするため、物理的な差分編集とは区別する。
- `pptx_get_job(jobId=latest)` は同じ利用者・会話の直近ジョブを状態にかかわらず返す。逐次修正の入力解決は成功済みVisual Deckだけを対象とするため、両者の `latest` の意味を混同しない。完了済みAnalyzeの大きな`result`は同一サーバー稼働中の最初の成功`pptx_wait_for_job`だけが返し、その後の`get_job`と同じ`wait`は本文を省略して既存結果から編集へ進む案内を返す。同じ利用者メッセージ・source・includeLayoutsのAnalyze再送は24時間同じjob IDへ束縛し、完全結果の誤認による新規解析を作らない。
- 非同期ジョブの通常フローは`pptx_wait_for_job`で最大45秒サーバー内待機し、`pptx_get_job`の短間隔反復でLibreChatの再帰上限とモデル出力枠を消費しない。`pptx_get_job`は障害復旧や待たない即時確認に使う。
- Visual Deckの検証失敗は `ToolValidationError` でコード・対象フィールド・修正指示をモデルへ返す。`PptxValidationException` をMCP境界から未処理のまま出し、モデルに同じ入力を推測再試行させない。
- 既存PPTXのanalyze、preview、replace、populate、create、refineでも受付時の`PptxValidationException`をMCP境界から漏らさず、`ToolValidationError`でcode、理由、再実行可否を返す。安全性検査で拒否した同一ファイルを推測再試行させない。
- `external_relationship`では「外部画像」「リンクされたExcel/Word等のOLE」「共有フォルダ・ネットワークドライブ・SharePoint上の別ファイル」が何を意味するかを利用者の言語で平易に説明する。通常のHTTP/HTTPSリンクは許可対象であることも明記し、拒否対象と混同させない。
- 添付PPTXの検索はLinux上でも`.pptx`と`.PPTX`を同じ拡張子として扱う。受付拒否時は、空・上限超過、ページ数、マクロ／ActiveX、破損・安全上限の違いを利用者の言語で説明し、修正版の再添付が必要かを明示する。
- 本番LibreChatは`X-LibreChat-Attachment-File-IDs`へ現在のリクエスト添付だけを渡す。`latest`と明示`file_id`はこのscope内で解決し、空scopeでは同じ利用者の過去・別会話の添付を再利用しない。ヘッダーを持たないローカル開発クライアントだけは従来のuser scopeへ後方互換フォールバックする。
- 翻訳など文字だけの既存PPTX更新は、`pptx_analyze(includeLayouts=false)`が返す全ページの文字入りshapeと編集可能なPowerPoint表のコンパクトな構造化結果から置換指定を作り、`pptx_replace_text`成功後にだけ全ページをプレビューする。解析応答はhostのtool結果圧縮で中間スライドを失わないよう、`slides`の`[slide_number, exact_texts[]]`タプルへ集約し、shape ID・theme・title・kindなど置換に不要な重複情報を含めない。表はセルごとに別のexact textとして投影し、見出しと本文を個別の翻訳対象にする。内部では`p:graphicFrame/a:tbl`の安定したshape IDで解析・置換し、置換指定はslide番号とexact textで行い、置換件数0の変更を成功扱いにしない。画像化された文字は編集可能テキストと誤認しない。元資料の文字起こし目的でプレビュー画像を先に取得し、モデルのコンテキストを消費しない。最初の置換後は同一メッセージの最新job IDをcheckpointとし、元資料から再開したり古いjobから分岐したりしない。同一メッセージで解析済みの元資料previewは例外なく拒否し、ユーザー意図をモデルが自己申告するoverrideを設けない。利用者が元資料の画像確認を明示した場合は、そのメッセージで解析より先にpreviewする。完成jobのpreviewは連続4ページずつ取得し、同じjobの同じページを再取得しない。既存プレースホルダーへ厳密に流し込むときだけ`includeLayouts=true`を使う。
- 置換が20件を超える場合は、`pptx_replace_text`を最大20件ずつ実行する。中間バッチは`isFinalBatch=false`とし、成功した`job_id`を次の`previousJobId`へ渡して変更を累積する。最後だけ`isFinalBatch=true`として全ページを描画する。
- `visual_draft_not_found`／`visual_draft_expired`／`visual_draft_not_editable`はterminal errorとして再試行を明示的に禁止する。汎用の「入力を直して同じtoolを再実行」案内へフォールバックさせない。
- PptxGenJSを含む外部レンダラーの生成物は、LibreOfficeで表示できてもPowerPoint互換とは限らない。生成直後と企業テンプレートへの合成後にOpenXmlValidatorを通し、新規検証エラーがある成果物は配布しない。
- `dom-to-pptx`のNode exporterはsandboxなしChromiumを起動するため、検証・再構築済みのoffline HTMLだけを入力し、非root、read-only、全capability削除、internal networkを維持する。runtimeでbrowserをdownloadせず、イメージへ固定インストールしたChromiumだけを使う。
- `dom-to-pptx`へ渡すflex内テキスト要素を内容幅のままにすると、PowerPointのtextbox内余白で行末1文字だけが折り返す。可変本文には`flex: 1; min-width: 0`で残り幅を与え、生成XMLのtextbox幅を確認する回帰テストを維持する。
- 互換rendererの色role、surface、style profileを拡張するときは`visual-v5`／`visual-v6-dom`へ限定し、保存済み`visual-v4` lineageの固定foreground、semantic tone、shape構成を変えない。`visual-v7-author-html`の視覚変更はモデルが渡すHTML/CSSで行い、server側へ新たな固定構図を増やさない。
- `VisualSlideSpec.speakerNotes`はvisible canvasとは別のPowerPoint発表者ノートである。`purpose`は1文、`talkScript`は発表用原稿として検証する。refineで省略されたノートは元ページから継承し、明示値だけを更新する。結果の`speaker_notes_count`で保持件数を返す。詳細はADR 0018を参照する。

## セキュリティと制約

- `rm`による削除は作業途中で繰り返さず、削除対象を記録して作業終盤に一覧を提示し、ユーザー承認後にまとめて実行する。承認前に実行しない。
- 入力上限は30MB、50スライド、ジョブ実行上限は10分、同時実行数は3とする。
- ZIP bomb、パストラバーサル、マクロ、ActiveX、外部リソースリレーション、非HTTP(S)ハイパーリンク、暗号化ファイルを拒否する。HTTP/HTTPSの標準ハイパーリンクは外部データを読み込まないため保持する。
- 成果物は生成から7日、または最初のダウンロードから24時間の早い方で削除する。
- MCP コンテナから外部ネットワークへ接続させない。調査や情報収集は LibreChat 側のモデルが担当する。
- ローカルCodex接続ではMCP本体の内部ネットワークを解除せず、`pptx-codex-proxy`だけを`127.0.0.1`へ公開する。本番LibreChatでは成果物専用プロキシを使い、`/mcp`を公開しない。
- 現行ホストでは `no-new-privileges` とDocker既定seccompの併用がruncのerrno 524で失敗する。既定seccompを維持するため同オプションは設定せず、非root、read-only、`cap_drop: ALL`で補完する。

## テスト方針

- 入力検証、認可境界、保持期限、PPTX パッケージ差分を優先してテストする。
- 編集前後で Open XML 検証と LibreOffice レンダリング確認を行う。
- SmartArt、グラフ、埋め込み Excel は各要素を含む代表ファイルを用意し、ゴールデンファイルテストを追加する。現在の提供テンプレートにはこれらの要素が含まれない。

## 注意点・落とし穴

- `p:ph` の `type` 属性は省略可能である。`PlaceholderShape.Type?.InnerText` を読み、省略時は `body` として扱う。存在しない型付き属性を `GetAttribute` で読むと Open XML SDK が例外を返す。
- Bedrock は PPTX 添付をモデル入力へ渡さない。LibreChat 上で `file_id` が会話に提示されない場合に備え、PPTX入力ツールの `sourceFileId` 省略時は呼出ユーザー配下の最新アップロードを解決する。明示された `file_id` は常に優先する。
- `pptx_analyze`の`layout_id`/`shape_id`/`placeholder_index`を作成・編集ツールへ直接コピーできるよう、対応するネスト入力キーもsnake_caseで維持する。`pptx_create_deck`は完成版の全ページを必須`slides`へまとめ、`sourceFileId`だけの呼出しを許容する説明にしない。
- 動作上必須の引数には`Required`属性を付け、MCPの公開JSON Schemaでも必須にする。ただしData Annotationsは実行時検証を代替せず、巨大な入力の生成自体も安定化しない。Visual Deckの新規生成は段階ドラフト、Visual Deckの自動修正は必須`revision`を持つ`pptx_refine_visual_slide`で1ページずつ逐次適用する。
- LibreChat v0.8.3-rc1 / `@librechat/agents` 3.1.51 はMCP画像artifactをBedrockへ再投入しない。LibreChat側のフェイルクローズなビルド時パッチを維持し、依存更新時に画像経路を再検証する。
- PptxGenJS 4.0.1は、PowerPointが修復を要求するOOXMLを生成することがある。`PptxGenJsOpenXmlNormalizer`でレンダラー所有のプレゼンテーションルート、表セル、グラフを正規化し、`node_modules`を直接改変しない。
- PptxGenJSへ負の`w`/`h`を渡すと`a:ext`へ負値をそのまま書き、Open XML検証に失敗する。右上・左上へ向かう線分は、正の幅・高さと`flipH`/`flipV`へ正規化して描画し、`PptxGenJsOpenXmlNormalizer`でも負のshape extentを位置移動＋反転属性へ補正する回帰テストを維持する。
- PptxGenJS 4.0.1が宣言する未使用の`image-size`は、parserを含まない`visual-renderer/vendor/image-size-disabled`へ固定し、呼出し時は即時失敗させる。`npm audit --omit=dev`を警告0件に保ち、shimのCommonJS／ESM拒否テストを維持する。画像登録toolのmetadataなしPNG化とrendererのhash・PNG signature・IHDR・宣言寸法検証も緩めない。
- LibreChatのmessage attachment画像は文書uploadと異なり`images/<user-id>/`へ保存される。導入時は`LibreChatImagesRoot`と`LibreChatUploadsRoot`を別々に読み取り専用mountし、実UI添付からresolverまでのE2Eを省略しない。
- 画像layoutを空欄や「ここに画像」のshapeで完成扱いしない。`Media`は同じ利用者・会話の有効なasset ID、alt text、cropを必須にし、素材がなければnative diagramまたは画像なしrecipeへ変更する。詳細は[ADR 0016](docs/adr/0016-conversation-scoped-image-assets-and-media-split.md)を参照する。
- 意味のある矢印・枠・吹き出し等は`pptx_prepare_visual_objects`で1回最大8件をまとめ、座標・生色・SVG/XML・URL・pathを入力に持たせない。1ページ最大3件、strong最大1件、会話最大24件を守り、Asset PlanへIDを固定する。補助図形は空のプレースホルダとして浮かせず、吹き出しは文字を内包し、枠は既存の焦点領域を囲み、図解の括弧は安全帯へ置くなど意味layout別anchorへ自動配置する。PPTX本体は編集可能なネイティブ図形とし、tool resultはJSON textだけを返す。SVG ImageContentはBedrock／Anthropicへ再送するとprovider errorになるため公開しない。詳細は[ADR 0017](docs/adr/0017-native-semantic-diagrams-and-visual-objects.md)を参照する。
- 複合Visual Objectは`recipe=directionalCue|growthPath|focusCorners|annotationPin|sectionRule|cycleCue`の正規tupleだけを受理し、`auto`は従来の単一図形を維持する。Agentはアウトライン後に一度だけ全ページを意味分類してbatch prepareし、最低使用数を設けない。複合recipeは原則subtle/standardとし、strongは`visualPurpose=emphasis`のfocusCornersだけに限定してBrand Profile確定後の手戻りを防ぐ。content間の方向キューへ重複labelを置かず、focusCornersは全面枠へ戻さない。annotationPinは折れ線Chartだけに使い、短いlabelと実在する1始まりcategory／series番号を必須にする。他recipeへanchor番号を渡さない。既存layoutの意味が十分ならobjectを追加しない。詳細は[ADR 0019](docs/adr/0019-curated-semantic-visual-object-recipes.md)を参照する。
- `variant`を公開するときは必ず専用描画分岐と件数条件を同時に実装し、受理値が`auto`へ黙って退化しない回帰テストを追加する。NativeDiagramのnode/edge上限とacyclic検証を緩めず、巨大schemaや座標修正ループを再導入しない。
- `addTable`の垂直中央揃えは`valign: "middle"`を使う。`"mid"`は表セルへ不正な`anchor="mid"`として出力されるため、正規化処理と回帰テストも維持する。
- PptxGenJSの棒グラフは系列内の`c:dPt`を`c:dLbls`より後ろへ出力する場合がある。Open XMLの要素順に合わせて個別データ点をラベルより前へ移し、実データ点を含む回帰テストを維持する。
- PptxGenJSはノート未指定ページにも空のノートスライドを生成する。`PptxGenJsOpenXmlNormalizer`は空のノートだけを削除し、非空ノートは正規化済みの単一ノートマスターへ接続する。企業テンプレートへの合成時も非空ノートを保持し、テンプレート側または生成側のノートマスターをページごとに複製しない。ノートはPPTX受領者が読めるため、秘密情報や内部思考を保存しない。
