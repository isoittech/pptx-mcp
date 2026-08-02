# LibreChat統合

## MCP接続

`integrations/librechat/librechat.fragment.yaml` を既存 `librechat.yaml` へマージします。LibreChat v0.8系では本番MCPに Streamable HTTP を使用し、`allowedAddresses` へコンテナ内の正確なホスト名とポートを追加します。

ユーザーIDと会話IDはLibreChatのヘッダープレースホルダーから渡します。これらをツール引数へ公開しないでください。共有秘密はLibreChatとpptx-mcpの環境変数に同じ値を設定します。

## アップロード連携

現在のアダプターはLibreChatローカルストレージの次の命名規則を読み取り専用マウントから解決します。

```text
/app/uploads/<user-id>/<file-id>__<original-name>.pptx
```

S3等へ移行する場合は `InputFileResolver` を、LibreChatの認証済み内部ファイル取得APIを呼ぶ実装へ交換します。MCPツールから任意URLを取得する方式にはしません。

## 成果物公開

`PPTX_MCP_PUBLIC_BASE_URL` は利用者のブラウザーから到達可能なHTTPS URLです。次の経路だけをリバースプロキシで公開します。

```text
GET /artifacts/{job_id}/{file_name}?token=...
```

`/mcp` はLibreChat内部ネットワークからだけ到達可能にします。署名付きURLをアクセスログへ残さない設定を推奨します。

## Claudeへの運用指示

- 資料編集前に `pptx_analyze` を実行する。
- 新規資料では解析結果の `layout_id` とプレースホルダー `shape_id` を使って `pptx_create_deck` を実行する。
- 編集対象が一意でなければ候補を列挙し、利用者の選択まで更新しない。
- ジョブ完了後は全スライドのプレビューを提示する。
- 自動視覚評価は行わず、利用者の指示を受けて会話的に再編集する。
- MCPに情報収集を依頼しない。外部情報はLibreChat側で収集し、構造化した内容だけを渡す。
