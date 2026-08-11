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
- 白紙生成は `VisualDeckSpec` を固定PptxGenJSレンダラーへ渡す。任意JavaScript、任意座標、URL、ローカルパスをツール入力へ追加しない。
- `VisualDeckSpec` は21種類の意味レイアウト、9テーマ、`design`、`variant`、組み込みアイコンを視覚語彙として提供する。固定構図へ内容を押し込まず、Opusが内容に合う構図を選べる語彙を増やす。
- 楽譜は`MusicScore`へ音高・音価・ウクレレ弦・フレット・指番号を渡し、PowerPointネイティブの線・図形・テキストで五線譜とTABを描く。任意座標を公開せず、音高と調弦・弦・フレットの一致を検証する。音楽記号は同梱したBravura 1.392の輪郭を`custGeom`へ変換し、手描き近似、画像貼付、閲覧側フォント依存へ戻さない。OTFとSIL OFL原文は必ず一緒に更新する。
- 文字量の多い説明は`StructuredBrief`の2〜3セクションへ分け、評価軸×選択肢は`Scorecard`の編集可能なPowerPoint表にする。`density=detailed`はフォント縮小だけで実装せず、外周余白、見出し領域、間隔、罫線、影を一体で切り替える。
- `tone` は意味語の別名またはRGB色を許容し、自然な表現を固定4語へ押し込めない。組み込みアイコンもモデルが実際に使う業務語彙をE2Eで確認して拡張し、未知値を黙って別アイコンへ置換しない。
- テンプレートのリストは `DeckField.paragraphs` / `TemplateField.paragraphs` を使い、1項目1段落で `Plain`、`Bullet`、`Numbered` と0〜4の `level` を指定する。本文へ `■`、`・`、`1.` 等を手入力しない。
- `pptx_analyze` の `theme` はaccent1〜3、light1、dark1、日本語優先の見出し・本文フォントを返す。白紙生成へ企業スタイルを移す場合はこの値を `VisualDeckSpec.theme` に使う。
- 新規Visual Deckは`pptx_start_visual_deck`、最大4ページずつの`pptx_add_visual_slides_to_draft`、finishの順で作る。最大50ページの完全仕様を一度に要求する公開ツールへ戻さない。`startSlideNumber`は省略可能とし、サーバーが受理済み末尾から追番する。明示値は順序検証にだけ使う。
- start時にテンプレート、テーマ、デザインを固定し、finishでの変更を拒否する。テンプレート抽出値は未指定テーマ項目だけを補完し、明示色・フォントを上書きしない。`pptx_create_deck` は既存プレースホルダー配置への厳密な流し込みが明示された場合だけ使う。
- 導入環境の既定テンプレートは外部マウントと`DefaultTemplateId`で指定し、実PPTX・会社名・ロゴ・固有文言をOSSへ含めない。起動時に検証・解析し、startの`templateSourceFileId=default`で自動適用する。添付テンプレートはstartで`latest`または`file_id`、テンプレートなしは`none`を選ぶ。
- Brand Profileは`BrandProfilesRoot/<profile-id>/brand-profile.json`から読み取り専用で起動時に検証し、会社固有値、URL、パスをOSSコードへ含めない。`pptx_get_design_catalog`は絞り込み式のcompact catalog、`pptx_validate_design_brief`はuser・conversation・profile version/hashへ束縛した期限付き`brief_id`を担当する。`RequireDesignBrief`はOSS既定offとし、導入環境で有効化した場合だけstart前に必須化する。詳細はADR 0013を参照する。
- Design Briefを使うdraftは全ページのAsset Planとlayout recipeを固定し、add時に`recipeId`、意味layout、密度、実装済みvariantを照合する。画像挿入未実装の段階では`userUpload`や`approvedLibrary`を完成素材として扱わず、nativeDrawまたは画像なしlayoutのfallbackと外部素材を必須としないrecipeを確定する。素材を使わない項目は`preferred_medium=none`、`acquisition=none`、`fallback=none`、`status=omitted`、`license_status=notRequired`の組合せにし、`noAssetLayout`は`userUpload`または`approvedLibrary`の`fallbackSelected`にだけ使う。
- 完成したprofile-bound deckには秘密を含まないDesign Brief監査snapshotとページrecipe契約を保存する。refineでは元ページの`recipeId`、kind、実効density、variantと監査snapshotを保持し、第1段階のinsertは新規Asset Planを検証できないため拒否する。期限付き`brief_id`はjob payloadへ保存しない。
- 生成・編集後は `pptx_get_preview_images` で全ページをClaudeへ渡し、問題時は宣言型仕様を最大2回まで修正する。通常の欠け・重なりに加え、見出しだけで話が追えるか、読む順序が一意か、1ブロック1論点か、強調色が概ね15%以内か、本文が9pt未満に見えないかを確認する。
- Bedrockから白紙資料とブランドVisual Deckを視覚修正する場合は、`pptx_refine_visual_slide` へ完全な差し替えページを1枚ずつ渡し、各成功後に `jobId=latest` で次ページを直す。ジョブへルート・親・修正巡を保存し、古い版からの分岐、一括ページ修正、3巡目をサーバー側で拒否する。成功後のstart再実行も拒否し、初回生成失敗時の全体再試行だけ1回許可する。
- 成功済みVisual Deckへページを増やす場合は`pptx_insert_visual_slides`へ追加ページだけを渡す。新規ドラフトを開始せず、既存ページも再送しない。サーバー側で元仕様へ挿入し、ブランド資料では元テンプレートとレイアウトを継承する。第1段階では完成版全体を再レンダリングするため、物理的な差分編集とは区別する。
- `pptx_get_job(jobId=latest)` は同じ利用者・会話の直近ジョブを状態にかかわらず返す。逐次修正の入力解決は成功済みVisual Deckだけを対象とするため、両者の `latest` の意味を混同しない。
- 非同期ジョブの通常フローは`pptx_wait_for_job`で最大45秒サーバー内待機し、`pptx_get_job`の短間隔反復でLibreChatの再帰上限を消費しない。`pptx_get_job`は障害復旧や待たない即時確認に使う。
- Visual Deckの検証失敗は `ToolValidationError` でコード・対象フィールド・修正指示をモデルへ返す。`PptxValidationException` をMCP境界から未処理のまま出し、モデルに同じ入力を推測再試行させない。
- PptxGenJSを含む外部レンダラーの生成物は、LibreOfficeで表示できてもPowerPoint互換とは限らない。生成直後と企業テンプレートへの合成後にOpenXmlValidatorを通し、新規検証エラーがある成果物は配布しない。
- rendererの色role、surface、style profileを拡張するときは`visual-v5`へ限定し、保存済み`visual-v4` lineageの固定foreground、semantic tone、shape構成を変えない。dark surface、明色role、v4代表XMLのNode回帰テストを維持する。

## セキュリティと制約

- `rm`による削除は作業途中で繰り返さず、削除対象を記録して作業終盤に一覧を提示し、ユーザー承認後にまとめて実行する。承認前に実行しない。
- 入力上限は30MB、50スライド、ジョブ実行上限は10分、同時実行数は3とする。
- ZIP bomb、パストラバーサル、マクロ、ActiveX、外部リレーション、暗号化ファイルを拒否する。
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
- PptxGenJS 4.0.1が依存する`image-size` 1.2.1には2026-08-08時点で修正版のないICNS/JXL/HEIFの無限ループDoSがある。Visual Deckレンダラーは画像入力と`addImage`を公開・使用しないため脆弱な解析経路へ到達しない状態をテストで固定し、修正版公開後に更新する。画像入力を追加する場合は先にこの依存を解消する。
- `addTable`の垂直中央揃えは`valign: "middle"`を使う。`"mid"`は表セルへ不正な`anchor="mid"`として出力されるため、正規化処理と回帰テストも維持する。
- PptxGenJSの棒グラフは系列内の`c:dPt`を`c:dLbls`より後ろへ出力する場合がある。Open XMLの要素順に合わせて個別データ点をラベルより前へ移し、実データ点を含む回帰テストを維持する。
- PptxGenJSは空のノートスライドとPowerPointが修復を要求するノートマスターを自動生成する。`VisualDeckSpec`は発表者ノートを扱わないため、白紙生成では`PptxGenJsOpenXmlNormalizer`がノート関連パーツを完全に削除し、企業テンプレートへの合成時も生成側のノートスライドを削除してテンプレート側のノートマスターだけを保持する。
