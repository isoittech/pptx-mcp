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
- 生成・編集後は `pptx_get_preview_images` で全ページをClaudeへ渡し、問題時は宣言型仕様を最大2回まで修正する。

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
- Bedrockが大きな`pptx_create_deck`を空引数で先行実行する場合は、スキーマ検証エラーではなく`input_required`を返して全`slides`付き再実行を促す。視覚修正では全仕様を再送せず、`pptx_refine_deck`へ変更ページだけを渡して元ジョブの仕様を再利用する。
- LibreChat v0.8.3-rc1 / `@librechat/agents` 3.1.51 はMCP画像artifactをBedrockへ再投入しない。LibreChat側のフェイルクローズなビルド時パッチを維持し、依存更新時に画像経路を再検証する。
