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
- `VisualDeckSpec` は17種類の意味レイアウト、9テーマ、`design`、`variant`、組み込みアイコンを視覚語彙として提供する。固定構図へ内容を押し込まず、Opusが内容に合う構図を選べる語彙を増やす。
- `tone` は意味語の別名またはRGB色を許容し、自然な表現を固定4語へ押し込めない。組み込みアイコンもモデルが実際に使う業務語彙をE2Eで確認して拡張し、未知値を黙って別アイコンへ置換しない。
- テンプレートのリストは `DeckField.paragraphs` / `TemplateField.paragraphs` を使い、1項目1段落で `Plain`、`Bullet`、`Numbered` と0〜4の `level` を指定する。本文へ `■`、`・`、`1.` 等を手入力しない。
- `pptx_analyze` の `theme` はaccent1〜3、light1、dark1、日本語優先の見出し・本文フォントを返す。白紙生成へ企業スタイルを移す場合はこの値を `VisualDeckSpec.theme` に使う。
- 企業テンプレートで華やかな新規資料を作る場合は `pptx_create_branded_visual_deck` を既定とする。プレースホルダー0個のレイアウトへVisual Deckを接続し、マスター、ロゴ、フッターと編集可能な図形・グラフを両立する。`pptx_create_deck` は既存プレースホルダー配置への厳密な流し込みが明示された場合だけ使う。
- 導入環境の既定テンプレートは外部マウントと`DefaultTemplateId`で指定し、実PPTX・会社名・ロゴ・固有文言をOSSへ含めない。起動時に検証・解析し、`pptx_create_visual_deck`は明示的に無効化されない限り既定テンプレートを自動適用する。添付した別テンプレートは明示`sourceFileId`でその処理だけ上書きする。
- 生成・編集後は `pptx_get_preview_images` で全ページをClaudeへ渡し、問題時は宣言型仕様を最大2回まで修正する。
- Bedrockから白紙資料とブランドVisual Deckを視覚修正する場合は、`pptx_refine_visual_slide` へ完全な差し替えページを1枚ずつ渡し、各成功後に `jobId=latest` で次ページを直す。修正は累積し、元テンプレートとレイアウトもジョブから再利用する。単一ページ修正が`Succeeded`を直接返した場合は`pptx_get_job`を重ねず、LibreChatの再帰上限を節約する。`pptx_refine_visual_deck` の一括配列は確実に構造化入力できるクライアント向けに残す。
- `pptx_get_job(jobId=latest)` は同じ利用者・会話の直近ジョブを状態にかかわらず返す。逐次修正の入力解決は成功済みVisual Deckだけを対象とするため、両者の `latest` の意味を混同しない。
- 非同期ジョブの通常フローは`pptx_wait_for_job`で最大45秒サーバー内待機し、`pptx_get_job`の短間隔反復でLibreChatの再帰上限を消費しない。`pptx_get_job`は障害復旧や待たない即時確認に使う。
- Visual Deckの検証失敗は `ToolValidationError` でコード・対象フィールド・修正指示をモデルへ返す。`PptxValidationException` をMCP境界から未処理のまま出し、モデルに同じ入力を推測再試行させない。
- PptxGenJSを含む外部レンダラーの生成物は、LibreOfficeで表示できてもPowerPoint互換とは限らない。生成直後と企業テンプレートへの合成後にOpenXmlValidatorを通し、新規検証エラーがある成果物は配布しない。

## セキュリティと制約

- 入力上限は30MB、50スライド、ジョブ実行上限は10分、同時実行数は3とする。
- ZIP bomb、パストラバーサル、マクロ、ActiveX、外部リレーション、暗号化ファイルを拒否する。
- 成果物は生成から7日、または最初のダウンロードから24時間の早い方で削除する。
- MCP コンテナから外部ネットワークへ接続させない。調査や情報収集は LibreChat 側のモデルが担当する。
- 現行ホストでは `no-new-privileges` とDocker既定seccompの併用がruncのerrno 524で失敗する。既定seccompを維持するため同オプションは設定せず、非root、read-only、`cap_drop: ALL`で補完する。

## テスト方針

- 入力検証、認可境界、保持期限、PPTX パッケージ差分を優先してテストする。
- 編集前後で Open XML 検証と LibreOffice レンダリング確認を行う。
- SmartArt、グラフ、埋め込み Excel は各要素を含む代表ファイルを用意し、ゴールデンファイルテストを追加する。現在の提供テンプレートにはこれらの要素が含まれない。

## 注意点・落とし穴

- `p:ph` の `type` 属性は省略可能である。`PlaceholderShape.Type?.InnerText` を読み、省略時は `body` として扱う。存在しない型付き属性を `GetAttribute` で読むと Open XML SDK が例外を返す。
- Bedrock は PPTX 添付をモデル入力へ渡さない。LibreChat 上で `file_id` が会話に提示されない場合に備え、PPTX入力ツールの `sourceFileId` 省略時は呼出ユーザー配下の最新アップロードを解決する。明示された `file_id` は常に優先する。
- `pptx_analyze`の`layout_id`/`shape_id`/`placeholder_index`を作成・編集ツールへ直接コピーできるよう、対応するネスト入力キーもsnake_caseで維持する。`pptx_create_deck`は完成版の全ページを必須`slides`へまとめ、`sourceFileId`だけの呼出しを許容する説明にしない。
- Bedrockが大きな作成ツールを空引数で先行実行する場合は、スキーマ検証エラーではなく`input_required`を返して完全入力での再実行を促す。Visual Deckの自動修正は大きな`revisions`配列を避け、必須`revision`を持つ`pptx_refine_visual_slide`で1ページずつ逐次適用する。
- LibreChat v0.8.3-rc1 / `@librechat/agents` 3.1.51 はMCP画像artifactをBedrockへ再投入しない。LibreChat側のフェイルクローズなビルド時パッチを維持し、依存更新時に画像経路を再検証する。
- PptxGenJS 4.0.1は、PowerPointが修復を要求するOOXMLを生成することがある。`PptxGenJsOpenXmlNormalizer`でレンダラー所有のプレゼンテーションルートとグラフだけを正規化し、`node_modules`を直接改変しない。
- PptxGenJSの棒グラフは系列内の`c:dPt`を`c:dLbls`より後ろへ出力する場合がある。Open XMLの要素順に合わせて個別データ点をラベルより前へ移し、実データ点を含む回帰テストを維持する。
- PptxGenJSのスライドを企業テンプレートへ1枚ずつ`AddPart`すると、空のノートスライド経由で共有ノートマスターがページごとに複製される。`VisualDeckSpec`は発表者ノートを扱わないため、合成時に生成側のノートスライドを削除し、企業テンプレート側のノートマスターだけを保持する。
