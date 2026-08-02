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
- LibreOffice と Poppler による全ページ PNG プレビュー
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
- `pptx_get_job`
- `pptx_cancel_job`

処理ツールはすぐに `job_id` を返します。Claude は `pptx_get_job` をポーリングし、完了後にプレビューとPPTXのリンクをユーザーへ提示します。

## テスト

ローカルに .NET SDK は不要です。

```bash
docker compose build test
docker compose run --rm test
```
