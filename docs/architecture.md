# アーキテクチャ

## 目的と境界

LibreChat/Claude が依頼理解、構成案、文章生成、必要な情報収集を担当し、pptx-mcp は決定済みの構造化指示を PPTX へ安全に反映します。MCP サーバー自身は外部インターネットへアクセスしません。

```text
利用者
  │ アップロード・会話・ダウンロード
  ▼
LibreChat / Claude
  │ Streamable HTTP MCP
  │ 信頼済み user/conversation headers
  ▼
pptx-mcp ──受付──▶ 永続ジョブキュー（最大3並列、最大10分）
  │                         │
  │ read-only               ├─ Open XML SDK: 検査・精密編集
  ▼                         ├─ 固定PptxGenJS: 宣言型の白紙・ブランド視覚生成
LibreChat uploads            ├─ Open XML SDK: Visual Deckと企業マスターの合成
                              └─ Poppler: PNG化
                              │
                              ├─ LibreOffice: PDF化
                              ▼
                       期限付き成果物ストレージ
                         ┌────┴───────────────┐
                         │                    │
                  MCP画像ブロック       署名付きHTTPS URL
                         │                    │
                         ▼                    ▼
                    Claude視覚評価          利用者
                         │ 問題時は仕様修正
                         └────最大2回再生成
```

## 処理モデル

MCP の通常応答には30MBのPPTXや最大50枚の画像を一括で埋め込みません。更新系ツールは `job_id` を返し、`pptx_get_job` が状態、解析JSON、短時間有効な成果物URLを返します。視覚評価時だけ `pptx_get_preview_images` が指定された1〜4枚をMCP画像ブロックとして返します。サービス再起動時は実行中ジョブを待機状態へ戻して再処理します。

## テンプレート戦略

企業テンプレートはテーマ、マスター、レイアウト、フォント、余白を正本とします。既存スライドは `slide_number + shape_id` で一意指定して更新できます。既存レイアウトへ厳密に内容を流し込む新規生成では、解析で得た `layout_id` とプレースホルダーの `shape_id` を使い、選択された定義済みレイアウトから1〜50枚のスライドを構成します。

華やかさとブランド保持を両立する場合は `pptx_create_branded_visual_deck` を使います。MCPが `pptx_analyze` と同じ処理でaccent色、背景色、文字色、見出し・本文フォントを自動抽出し、Visual Deckのテーマへ適用します。固定PptxGenJSレンダラーで編集可能な図形・グラフを生成した後、Open XMLでテンプレートの既存スライドだけを除去し、生成スライドをプレースホルダー0個の白紙レイアウトへ接続します。`templateLayoutId=auto` は「白紙（フッター有）」を優先し、写真背景は避けます。テンプレートのマスター、ロゴ、フッター、ページ設定は成果物側に残ります。

テンプレートとVisual Deckの縦横比が一致しない場合や、白紙レイアウトがない場合は合成を拒否します。Visual Deck側の独自フッターは合成時だけ抑制し、企業フッターとの二重表示を防ぎます。詳細な判断は [ADR 0004](adr/0004-branded-visual-composition.md) に記録します。

AIにJavaScriptやOpen XMLを直接生成・実行させません。Claudeから受けるのはJSON Schemaで制約したスライド仕様・編集命令だけです。通常文は `text`、箇条書き・番号付き手順は項目単位の `paragraphs` として受け取り、DrawingMLの実段落、箇条書き、自動採番へ変換します。

## 白紙からの視覚的な生成

テンプレートがない場合は `pptx_create_visual_deck` を使います。Claudeはスライドの意味に応じて次の固定レイアウトを選びます。

- title、agenda、section、statement、bullets
- cards、metrics、comparison、process、timeline
- matrix、funnel、roadmap、chart、dashboard
- quote、closing

テーマは `midnight`、`aurora`、`sunset`、`forest`、`minimal`、`ocean`、`berry`、`clay`、`cyber` の9種です。色とフォントは検証済みの範囲で上書きできます。Opusは `design.style`、`density`、`motif` と `variant` を使って、同じ意味レイアウトでも資料固有の視覚表現を選びます。固定PptxGenJSレンダラーはテキスト、図形、組み込みアイコン、テーマ色、編集可能グラフ、グラフ用埋め込みワークブックを生成します。入力にファイルパス、URL、画像、JavaScript、任意座標を持たせないため、表現力を上げてもコード実行境界は広げません。

固定レイアウトはモデルのデザイン判断を置き換えるものではなく、安全に実行できる視覚語彙です。モデルがストーリー、強調対象、構図、視覚モチーフを決め、レンダラーは整列、最小余白、編集可能性、ファイル整合性を保証します。6枚以上で構図が4種類未満、同一構図が3枚連続、文字中心ページが過半数の場合は `design_warnings` を返します。

## 自動視覚リフレクション

MCPサーバー指示とツール説明に次のエージェントループを定義します。

1. 生成・編集ジョブを完了させる。
2. 全スライドを1〜4枚ずつMCP画像ブロックとして取得する。
3. Claudeが文字切れ、はみ出し、重なり、文字サイズ、余白、整列、コントラスト、情報階層、密度、バランス、全体一貫性を確認する。
4. 問題があれば厳密なプレースホルダー資料は `pptx_refine_deck`、白紙資料とブランドVisual Deckは `pptx_refine_visual_slide` へ完全な差し替えページを1枚ずつ渡して再生成する。
5. 最大2回で収束させ、視覚確認後にダウンロードリンクを提示する。

`pptx_refine_visual_slide` は `jobId=latest` で同じ利用者・会話の最新成功Visual Deckだけを解決します。複数ページは各ジョブの成功後に1枚ずつ適用するため、修正済み仕様と企業テンプレートが次のジョブへ累積します。Bedrockが大きな `revisions` 配列を省略する問題を避けつつ、会話境界を越えたジョブ参照は許しません。単一ページ修正は最大30秒だけサーバー内で完了を待ち、最終状態を返せた場合は追加ポーリングを不要にします。一括入力が安定したクライアント向けには `pptx_refine_visual_deck` も維持します。

これはMCPサーバーが別のモデルAPIを直接呼ぶループではなく、LibreChat上のClaudeがツール呼び出しを継続するエージェント駆動方式です。モデル認証情報をMCPへ持ち込まず、会話文脈を保ったまま評価できます。

## 編集エンジン

`IPresentationEngine` により以下を分離します。

- `OpenXmlPresentationEngine`: OOXMLパーツを限定的に変更し、未対象要素を保持する既定実装
- `AsposePresentationEngine`（候補）: SmartArtノード操作や描画互換性を商用ライセンス込みで評価する差し替え実装

既存ファイル更新では変更対象パーツを記録し、編集前後の Open XML 検証エラー数を比較します。LibreOfficeで全ページを描画できることも完了条件です。

## 曖昧な編集対象

`pptx_analyze` はスライド番号、シェイプID、シェイプ名、種類、テキストを返します。候補が複数ある場合、Claudeは更新を実行せず、候補とプレビューを利用者へ示して選択を求めます。選択後はスライド番号とシェイプIDを編集命令へ固定します。提供テンプレートでは同一スライド内のシェイプ名重複が確認されているため、シェイプ名だけを永続キーにしません。

## セキュリティ

- LibreChatとのMCP通信は共有秘密で認証し、ユーザーID・会話IDはLibreChatが挿入するヘッダーだけを信頼します。
- ツール引数でローカルパス、ユーザーID、出力URLを指定できません。
- `file_id` は該当ユーザーのアップロードディレクトリ内だけで解決し、作業領域へコピーします。Bedrock経由で`file_id`がモデルに提示されない場合、`sourceFileId`省略時に同じユーザー境界内の最新PPTXを選択します。
- マクロ、ActiveX、外部リレーション、暗号化PPTX、ZIP bombを拒否します。
- 成果物URLはファイル名と期限をHMAC署名した capability URL です。
- MCP用内部ネットワークは `internal: true` とし、成果物配信だけを認証付きリバースプロキシ経由で公開します。
- コンテナは非root、read-only、全Linux capability削除、Docker既定seccompで実行します。現行ホストでは `no-new-privileges` と既定seccompの併用がruncのerrno 524で起動失敗するため、seccompを無効化せず同オプションだけを省いています。
- ログに本文、ファイル内容、秘密値、署名付きURLを記録しません。

## 保持期間

ジョブ作成時に `created_at + 7日` を設定します。成果物PPTXへの最初の有効なGETで `first_downloaded_at` を原子的に設定し、実効削除期限を次の早い方とします。

```text
min(created_at + 7日, first_downloaded_at + 24時間)
```

プレビュー画像の閲覧だけでは24時間タイマーを開始しません。

## 未実装項目の受け入れ条件

- SmartArt: テキスト、ノード追加・削除を行い、種類は変更しない。
- グラフ: データ、カテゴリ、タイトル、系列追加・削除、凡例、軸、色を更新し、種類は変更しない。グラフXMLのキャッシュと埋め込みブックを同期する。
- 埋め込みExcel: 値、数式、範囲、行列追加・削除を対象とし、シート追加、ピボット、外部接続は対象外とする。
- 各操作は提供テンプレートと代表サンプルでPowerPoint/LibreOfficeのゴールデンテストを通す。

## 参考資料

- [LibreChat MCP configuration](https://www.librechat.ai/docs/configuration/librechat_yaml/object_structure/mcp_servers)
- [Model Context Protocol C# SDK](https://github.com/modelcontextprotocol/csharp-sdk)
- [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
- [PptxGenJS](https://github.com/gitbrent/PptxGenJS)
- [PresentationML document structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document)
- [LibreOffice command-line parameters](https://help.libreoffice.org/latest/en-GB/text/shared/guide/start_parameters.html)
- [参考記事: Claude Opus 4.6 の PowerPoint 生成手法](https://zenn.dev/microsoft/articles/how-the-claude-opus46-generate-pptx)
