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
- 白紙から意味ベースの20レイアウト、9テーマ、デザイン方針、構図バリエーションを使って視覚的な16:9デッキを生成
- 文字量の多い説明を2〜3列へ構造化する `structuredBrief` と、評価軸×選択肢を編集可能なPowerPoint表にする `scorecard`
- `density=detailed` で外周余白、見出し領域、罫線、カード影、内部間隔を一体的に切り替える高密度デザイン
- 企業テンプレートのマスター、ロゴ、フッター、ページ設定を保ち、テーマを自動抽出して意味ベースのVisual Deckを合成
- 外部登録された既定テンプレートの起動時検証・事前解析と、新規資料への自動適用
- 導入環境から任意に注入できる初回assistant案内
- PptxGenJSによる編集可能なグラフと埋め込みデータブックの生成
- `musicScore`による編集可能な五線譜、Bravura由来の高品質な音楽記号、小節線、ウクレレTAB、指番号の色分け
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

ローカルCodexから接続する場合、通常の`compose.yaml`は`pptx-codex-proxy`を`127.0.0.1:${PPTX_MCP_PORT:-18080}`へ公開します。MCP本体は外部通信できない内部ネットワークに残り、ローカル入口だけが`/mcp`、署名付き`/artifacts/...`、readinessを中継します。Codex側ではBearer認証に加え、`X-LibreChat-User-ID`と`X-LibreChat-Conversation-ID`へローカル用の固定識別子を設定してください。本番LibreChat統合では成果物専用プロキシを使い、`/mcp`を公開しません。

`PPTX_MCP_PUBLIC_BASE_URL` は利用者のブラウザーから到達できる HTTPS URL にしてください。MCP の内部URL `http://pptx-mcp:8080/mcp` とは別です。

## 既定テンプレート

既定テンプレートを使う場合は、PPTXをホスト側の非公開ディレクトリへ置き、読み取り専用でコンテナの `/data/pptx-templates` へマウントします。ファイル名は `<template-id>.pptx` とし、次を設定します。

```dotenv
PPTX_MCP_TEMPLATES_PATH=/absolute/path/to/pptx-templates
PPTX_MCP_DEFAULT_TEMPLATE_ID=organization-default
```

テンプレートIDは英数字、ハイフン、アンダースコアだけを使用できます。実ファイル、会社名、ロゴ、社内向け文言はOSSリポジトリやコンテナイメージへ含めません。設定済みの既定テンプレートは起動時にPPTX安全性検査と構造解析を行い、不正または欠落していれば起動を失敗させます。解析結果は内容のSHA-256単位で再利用するため、通常の新規生成前に `pptx_analyze` は不要です。

通常の新規資料は `pptx_start_visual_deck` で概要、完成ページ数、テンプレート、テーマ、デザインを確定し、`pptx_add_visual_slides_to_draft` で最大4ページずつ追加した後、`pptx_finish_visual_deck` で生成します。`templateSourceFileId` の既定値は `default` です。テンプレート不要は `none`、添付した別テンプレートは解析後に `latest` または `file_id` をstartへ指定します。選択内容はstart成功時に固定され、finishでは変更しません。

導入環境固有の初回案内が必要な場合は、任意の文言を次へ設定できます。

```dotenv
PPTX_MCP_FIRST_ASSISTANT_NOTICE=PowerPoint資料には導入環境の既定テンプレートを使用します。
```

この値はMCPサーバー指示へ追加され、会話内でまだ案内されていない場合に、最初のユーザー可視assistant応答の冒頭へ一度だけ表示するようモデルへ指示します。文言は導入環境で管理し、OSS側の既定値は空です。

## MCPツール

- `pptx_get_capabilities`
- `pptx_analyze`
- `pptx_render_preview`
- `pptx_replace_text`
- `pptx_populate_template`
- `pptx_create_deck`
- `pptx_refine_deck`
- `pptx_start_visual_deck`
- `pptx_add_visual_slides_to_draft`
- `pptx_finish_visual_deck`
- `pptx_finish_branded_visual_deck`
- `pptx_insert_visual_slides`
- `pptx_refine_visual_slide`
- `pptx_refine_visual_deck`
- `pptx_get_job`
- `pptx_wait_for_job`
- `pptx_get_preview_images`
- `pptx_cancel_job`

処理ツールはすぐに `job_id` を返します。Claude は `pptx_wait_for_job` を1回呼んでMCPサーバー内で最大45秒待ち、完了後に `pptx_get_preview_images` で全ページを1〜4枚ずつ実際に確認します。指定時間内に終わらない場合だけ、同じ`job_id`でもう一度待機します。文字切れ、重なり、可読性、整列、余白、コントラスト、情報密度、一貫性を確認し、問題ページだけを1枚ずつ最大2巡まで修正します。修正ラウンド、最新ジョブの直列性、上限はサーバー側でも強制し、成功後の全体再作成へ戻りません。

新規生成の段階ワークフローはAIが任意のJavaScriptや座標を実行する方式ではありません。AIは `statement`、`cards`、`metrics`、`comparison`、`structuredBrief`、`scorecard`、`musicScore`、`process`、`timeline`、`matrix`、`funnel`、`roadmap`、`chart`、`dashboard` などの意味ベースのレイアウトを選びます。さらに `design.style`、`density`、`motif` とスライド単位の `variant` で、Opusが資料固有のアートディレクションと構図を指定し、固定レンダラーが編集可能なPowerPoint要素へ変換します。既定テンプレートが登録されていれば企業マスターへ自動合成し、`templateSourceFileId=none` の場合だけ白紙生成になります。

長い説明を1つの本文枠へ押し込まず、見出しだけでも要点を追える2〜3個の `sections` に分ける場合は `structuredBrief` を使います。セクション合計は900文字までです。3セクション・600文字未満では上段の主論点と下段2論点からなるモザイクへ自動変更し、短い内容ほど本文と箇条書きを拡大します。600文字以上では3列を使い、情報量に合わない大きな空箱を避けます。複数案を評価軸で比べる場合は `scorecard` を使い、2〜4個の `options` と2〜6行の `criteria` を指定します。各評価セルは短い判定、根拠、意味色を持ち、成果物では編集可能なPowerPointネイティブ表になります。文字量が多い資料では `design.density=detailed` を指定すると、フォントだけを縮小せず、余白、タイトル領域、内部間隔、細い罫線、影なしカードをまとめて切り替えます。

`musicScore`は1〜8小節、合計64イベントまでの五線譜とウクレレTABを上下に併記します。各イベントへ`duration`、各音へ科学的音高の`pitch`、1=A/2=E/3=C/4=Gの`string`、`fret`、任意の`finger`を指定します。`tuning`は`high-g`または`low-g`です。音高と弦・フレットの不一致は入力検証で拒否します。ト音記号、拍子数字、符頭、旗、付点、休符、臨時記号はSIL Open Font LicenseのBravura 1.392から輪郭を取得し、画像やフォントではなくPowerPointカスタム図形として生成します。その他の五線、符幹、小節線、TAB線、フレット番号、指色マーカーも個別編集できるPowerPointネイティブ要素です。PowerPoint自体に楽譜の意味モデルはないため、移調やリズム変更に伴う自動再配置は行いません。Bravuraの原本ライセンスは`visual-renderer/assets/bravura/LICENSE.txt`に同梱しています。

新規Visual Deckは、完成ページ数とクリエイティブ方針を登録するstart、連続した1〜4ページを渡すadd、ドラフトIDだけで生成するfinishへ分割しています。addの`startSlideNumber`は省略でき、サーバーが受理済み末尾から自動計算します。ドラフトは利用者と会話の境界内だけで参照でき、1時間で失効します。成功済みデッキがある会話では通常のstartを拒否し、初回生成が失敗した場合の全体再試行も1回に制限します。ユーザーが別資料を明示的に求めた場合だけ`userRequestedNewWorkflow=true`で新しいワークフローを開始できます。

企業テンプレートを使いつつ同じ視覚表現が必要な場合は、テンプレートをstart時に選び、ドラフト完成後に `pptx_finish_branded_visual_deck` を使います。テンプレートのテーマ色と日本語フォントを自動抽出し、未指定のテーマ項目だけを補完して、プレースホルダーのない白紙レイアウトへ各スライドを接続します。startで明示した色とフォントはテンプレート抽出値より優先されます。これにより企業マスターのロゴ・フッターと、資料固有の配色、カード、工程、マトリクス、編集可能グラフ等を両立します。

メトリクスとカードの `tone` は `positive`、`critical`、`negative`、`info` 等の意味語または任意の `#RRGGBB` を受け付けます。カードは `search`、`compliance`、`decision`、`network`、`recovery` 等を含む編集可能な組み込みアイコンを利用できます。カスタムテーマで背景と文字のコントラストが不足する場合は、レンダラーが可読色へ自動補正します。

Visual Deckの入力検証エラーは `status=invalid_input`、エラーコード、対象フィールドを構造化して返します。モデルは推測で同じ呼び出しを繰り返さず、指摘されたフィールドだけを直せます。Closingの提言はPowerPointネイティブの箇条書きとして描画されます。

6枚以上の資料で構図が4種類未満、同じ構図が3枚連続、または文字中心のページが過半数になると、ジョブ結果の `design_warnings` に改善案を返します。視覚確認後は `pptx_refine_visual_slide` へ問題ページを1枚ずつ渡します。`jobId=latest` が直前の成功ジョブを選ぶため、複数ページの修正は逐次累積します。各ジョブはルート、親、修正巡、同巡の修正ページを記録し、古いジョブからの分岐、一括ページ修正、3巡目を拒否します。`pptx_refine_visual_deck`も互換用に残しますが、同じ1ページ制約を適用します。

成功済みVisual Deckへページを追加する場合は `pptx_insert_visual_slides` を使います。`slides` には追加分だけを渡し、既存ページを再送しません。`jobId` の既定値は `latest`、`afterSlideNumber` の省略時は末尾、`0` は先頭、正の値はそのページの直後へ挿入します。サーバーが元ジョブのタイトル、テーマ、デザイン、既存ページ、企業テンプレートとレイアウトを結合し、最大50ページの完成版を再生成します。通常はサーバー内で最大30秒待って最終状態を返します。この方式はAIの入力を追加ページ分に限定しますが、レンダラーとプレビュー生成は現段階では完成版全体を処理します。

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

2026-08-08時点ではPptxGenJSの推移依存`image-size`に、修正版未公開のICNS/JXL/HEIF解析DoSが報告されます。Visual Deckは画像、URL、ローカルパスを入力として受け取らず、レンダラーも`addImage`を使用しないため該当パーサーへ到達しません。この境界はNodeテストで固定しています。修正版が公開されたらロックファイルを更新してください。
