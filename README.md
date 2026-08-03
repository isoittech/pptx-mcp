# pptx-mcp

LibreChat 上の Claude から、PowerPoint 資料の解析、テンプレートへの流し込み、既存資料の更新、全ページプレビュー、ダウンロードを行う MCP サーバーです。

## 現在の実装範囲

- Streamable HTTP の MCP サーバー
- 30MB・50スライドまでの PPTX 入力検証
- ZIP bomb、パストラバーサル、マクロ、ActiveX、外部参照の拒否
- 最大3並列・10分タイムアウトの永続化された非同期ジョブ
- スライド、シェイプ、SmartArt、グラフ、埋め込みExcel有無の解析
- 通常シェイプと SmartArt 内テキストの置換
- 名前付きシェイプを持つ企業テンプレートへのテキスト流し込み
- 企業テンプレートの定義済みレイアウトから1〜50枚の新規デッキを生成
- 白紙から意味ベースの11レイアウトと5テーマを使って視覚的な16:9デッキを生成
- PptxGenJSによる編集可能なグラフと埋め込みデータブックの生成
- LibreOffice と Poppler による全ページ PNG プレビュー
- プレビュー画像をClaudeへ返し、最大2回まで自律修正する視覚リフレクション
- 15分有効の署名付き成果物URL
- 生成から7日、または最初のPPTXダウンロードから24時間の早い方で削除

SmartArtノード、グラフデータ、埋め込みExcel、既存デッキのスライド構成変更は、各要素を含む代表ファイルを使った互換性スパイク後に実装します。詳細は [architecture.md](docs/architecture.md) と [template-spike.md](docs/template-spike.md) を参照してください。

## 起動

1. `.env.example` を `.env` にコピーし、共有秘密は24文字以上、署名鍵は32文字以上の独立したランダム値に置き換え、LibreChat のアップロードディレクトリを設定します。
2. `docker compose build` を実行します。
3. `docker compose up -d pptx-mcp` を実行します。
4. `integrations/librechat/librechat.fragment.yaml` の内容を LibreChat 設定へ統合します。

`PPTX_MCP_PUBLIC_BASE_URL` は利用者のブラウザーから到達できる HTTPS URL にしてください。MCP の内部URL `http://pptx-mcp:8080/mcp` とは別です。

## MCPツール

- `pptx_get_capabilities`
- `pptx_analyze`
- `pptx_render_preview`
- `pptx_replace_text`
- `pptx_populate_template`
- `pptx_create_deck`
- `pptx_create_visual_deck`
- `pptx_get_job`
- `pptx_get_preview_images`
- `pptx_cancel_job`

処理ツールはすぐに `job_id` を返します。Claude は `pptx_get_job` をポーリングし、完了後に `pptx_get_preview_images` で全ページを1〜4枚ずつ実際に確認します。文字切れ、重なり、可読性、整列、余白、コントラスト、情報密度、一貫性に問題があれば宣言型仕様を修正して最大2回まで再生成し、その後にPPTXリンクを提示します。

白紙生成の `pptx_create_visual_deck` はAIが任意のJavaScriptや座標を実行する方式ではありません。AIは `title`、`metrics`、`comparison`、`process`、`timeline`、`chart` などの意味ベースのレイアウトと制限付きコンテンツをJSONで指定し、固定レンダラーが配置します。

## テスト

ローカルに .NET SDK は不要です。

```bash
docker compose build test
docker compose run --rm test
```

ビジュアルレンダラーの依存監査は次で実行します。

```bash
cd visual-renderer
npm audit --omit=dev
```
