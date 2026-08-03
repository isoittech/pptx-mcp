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
  ▼                         ├─ LibreOffice: PDF化
LibreChat uploads            └─ Poppler: PNG化
                              │
                              ▼
                       期限付き成果物ストレージ
                              │ 署名付きHTTPS URL
                              ▼
                           利用者
```

## 処理モデル

MCP の HTTP 応答には30MBのPPTXや最大50枚の画像を埋め込みません。更新系ツールは `job_id` を返し、`pptx_get_job` が状態、解析JSON、短時間有効な成果物URLを返します。サービス再起動時は実行中ジョブを待機状態へ戻して再処理します。

## テンプレート戦略

企業テンプレートはテーマ、マスター、レイアウト、フォント、余白を正本とします。既存スライドは `slide_number + shape_id` で一意指定して更新できます。新規生成では解析で得た `layout_id` とプレースホルダーの `shape_id` を使い、選択された定義済みレイアウトから1〜50枚のスライドを構成します。

AIにJavaScriptやOpen XMLを直接生成・実行させません。Claudeから受けるのはJSON Schemaで制約したスライド仕様・編集命令だけです。

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
- [PresentationML document structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document)
- [LibreOffice command-line parameters](https://help.libreoffice.org/latest/en-GB/text/shared/guide/start_parameters.html)
- [参考記事: Claude Opus 4.6 の PowerPoint 生成手法](https://zenn.dev/microsoft/articles/how-the-claude-opus46-generate-pptx)
