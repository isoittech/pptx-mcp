# pptx-mcp

LibreChat 上の Claude から、PowerPoint 資料の解析、テンプレートへの流し込み、既存資料の更新、全ページプレビュー、ダウンロードを行う MCP サーバーです。

## 現在の実装範囲

- Streamable HTTP の MCP サーバー
- 30MB・50スライドまでの PPTX 入力検証
- ZIP bomb、パストラバーサル、マクロ、ActiveX、外部参照の拒否
- 最大3並列・10分タイムアウトの永続化された非同期ジョブ
- スライド、シェイプ、テーマ色、日本語フォント、SmartArt、グラフ、埋め込みExcel有無の解析
- 通常シェイプと SmartArt 内テキストの置換
- 名前付きシェイプを持つ企業テンプレートへのテキスト流し込み
- 企業テンプレート内で編集可能な実箇条書き・自動採番・0〜4段階のインデントを生成
- 企業テンプレートの定義済みレイアウトから1〜50枚の新規デッキを生成
- 白紙から意味ベースの17レイアウト、9テーマ、デザイン方針、構図バリエーションを使って視覚的な16:9デッキを生成
- 企業テンプレートのマスター、ロゴ、フッター、ページ設定を保ち、テーマを自動抽出して意味ベースのVisual Deckを合成
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
- `pptx_refine_deck`
- `pptx_create_visual_deck`
- `pptx_create_branded_visual_deck`
- `pptx_refine_visual_slide`
- `pptx_refine_visual_deck`
- `pptx_get_job`
- `pptx_get_preview_images`
- `pptx_cancel_job`

処理ツールはすぐに `job_id` を返します。Claude は `pptx_get_job` をポーリングし、完了後に `pptx_get_preview_images` で全ページを1〜4枚ずつ実際に確認します。`pptx_get_job` の `jobId=latest` は同じ利用者・会話にある直近ジョブ（待機中・実行中を含む）を選びます。文字切れ、重なり、可読性、整列、余白、コントラスト、情報密度、一貫性に問題があれば宣言型仕様を修正して最大2回まで再生成し、その後にPPTXリンクを提示します。

白紙生成の `pptx_create_visual_deck` はAIが任意のJavaScriptや座標を実行する方式ではありません。AIは `statement`、`cards`、`metrics`、`comparison`、`process`、`timeline`、`matrix`、`funnel`、`roadmap`、`chart`、`dashboard` などの意味ベースのレイアウトを選びます。さらに `design.style`、`density`、`motif` とスライド単位の `variant` で、Opusが資料固有のアートディレクションと構図を指定し、固定レンダラーが編集可能なPowerPoint要素へ変換します。

企業テンプレートを使いつつ同じ視覚表現が必要な場合は `pptx_create_branded_visual_deck` を使います。テンプレートのテーマ色と日本語フォントを自動抽出し、Visual Deckを生成した後、プレースホルダーのない白紙レイアウトへ各スライドを接続します。これにより企業マスターのロゴ・フッターと、カード、工程、マトリクス、編集可能グラフ等を両立します。既存プレースホルダーへの厳密な流し込みが目的の場合だけ `pptx_create_deck` を使います。

メトリクスとカードの `tone` は `positive`、`critical`、`negative`、`info` 等の意味語または任意の `#RRGGBB` を受け付けます。カードは `search`、`compliance`、`decision`、`network`、`recovery` 等を含む編集可能な組み込みアイコンを利用できます。カスタムテーマで背景と文字のコントラストが不足する場合は、レンダラーが可読色へ自動補正します。

Visual Deckの入力検証エラーは `status=invalid_input`、エラーコード、対象フィールドを構造化して返します。モデルは推測で同じ呼び出しを繰り返さず、指摘されたフィールドだけを直せます。Closingの提言はPowerPointネイティブの箇条書きとして描画されます。

6枚以上の資料で構図が4種類未満、同じ構図が3枚連続、または文字中心のページが過半数になると、ジョブ結果の `design_warnings` に改善案を返します。単独のMetricsスライドは最大6指標を3列×2段で配置できます。視覚確認後、Bedrock/Claudeでは `pptx_refine_visual_slide` へ問題ページを1枚ずつ渡します。`jobId=latest` が同じ会話の直前の成功ジョブを選ぶため、複数ページの修正は逐次累積し、大きな一括入力を避けられます。単一ページ修正はサーバー内で最大30秒完了を待ち、通常は最終状態を直接返すため、成功時の追加ポーリングとLibreChatのグラフステップ消費も抑えます。`pptx_refine_visual_deck` は一括配列を確実に送信できるクライアント向けです。

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
