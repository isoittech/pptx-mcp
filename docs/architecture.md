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
                         │ 問題ページだけ仕様修正
                         └────最大2巡の差分再生成
```

## 処理モデル

MCP の通常応答には30MBのPPTXや最大50枚の画像を一括で埋め込みません。更新系ツールは `job_id` を返し、`pptx_wait_for_job` がサーバー内で最大45秒待って状態、解析JSON、短時間有効な成果物URLを返します。`pptx_get_job`は待たない即時確認用です。視覚評価時だけ `pptx_get_preview_images` が指定された1〜4枚をMCP画像ブロックとして返します。サービス再起動時は実行中ジョブを待機状態へ戻して再処理します。

## テンプレート戦略

企業テンプレートはテーマ、マスター、レイアウト、フォント、余白を正本とします。既存スライドは `slide_number + shape_id` で一意指定して更新できます。既存レイアウトへ厳密に内容を流し込む新規生成では、解析で得た `layout_id` とプレースホルダーの `shape_id` を使い、選択された定義済みレイアウトから1〜50枚のスライドを構成します。

OSS本体は企業固有テンプレートを保持せず、外部マウントされた `<template-id>.pptx` と導入環境の `DefaultTemplateId` だけを扱います。既定テンプレートは起動時に安全性検査と解析を完了し、欠落・不正時はfail-closedで起動を失敗させます。解析はPPTX内容のSHA-256単位でメモリキャッシュし、ジョブ領域へコピーした同一ファイルや同じ添付ファイルの再解析を避けます。

新規Visual Deckは、`pptx_start_visual_deck`で概要、完成ページ数、テンプレート、テーマ、デザインを固定し、`pptx_add_visual_slides_to_draft`で最大4ページずつ仕様を蓄積して、finishツールでジョブ化します。テンプレートは`default`、`none`、`latest`または添付`file_id`からstart時に選択し、finishでの変更を拒否します。既定テンプレートの判断は [ADR 0006](adr/0006-deployment-default-template.md)、段階入力は [ADR 0009](adr/0009-staged-visual-deck-input.md)、生成・修正ループの停止条件は [ADR 0012](adr/0012-visual-deck-workflow-guardrails.md) に記録します。

華やかさとブランド保持を両立する場合は、startで企業テンプレートを選び、ドラフト完成後に `pptx_finish_branded_visual_deck` を使います。MCPがaccent色、背景色、文字色、見出し・本文フォントを自動抽出し、startで未指定のテーマ項目だけを補完します。明示色とfontFaceは抽出値より優先されます。固定PptxGenJSレンダラーで編集可能な図形・グラフを生成した後、生成スライドをプレースホルダー0個の白紙レイアウトへ接続します。

テンプレートとVisual Deckの縦横比が一致しない場合や、白紙レイアウトがない場合は合成を拒否します。Visual Deck側の独自フッターは合成時だけ抑制し、企業フッターとの二重表示を防ぎます。詳細な判断は [ADR 0004](adr/0004-branded-visual-composition.md) に記録します。

AIにJavaScriptやOpen XMLを直接生成・実行させません。Claudeから受けるのはJSON Schemaで制約したスライド仕様・編集命令だけです。通常文は `text`、箇条書き・番号付き手順は項目単位の `paragraphs` として受け取り、DrawingMLの実段落、箇条書き、自動採番へ変換します。

## 白紙からの視覚的な生成

新規Visual Deckの段階ワークフローでは、Claudeがスライドの意味に応じて次の固定レイアウトを選びます。既定テンプレートがない場合、または利用者が明示的に不要とした場合は白紙へ描画します。

最大50ページの完全な`VisualDeckSpec`を1回のツール入力で生成させると、公開JSON Schemaで`deck`を必須にしてもBedrock Claude Opus 5が空呼び出しを先行することをE2Eで確認しました。このため一括作成ツールは公開せず、概要、最大4ページの連続バッチ、生成確定へ分割します。ドラフトは利用者と会話で分離し、1時間で失効します。addの`startSlideNumber`は任意とし、省略時は受理済み末尾からサーバーが算出します。明示値がある場合は正しい次番号との一致を検証します。

- title、agenda、section、statement、bullets
- cards、metrics、comparison、structuredBrief、scorecard、dataTable、media、nativeDiagram、musicScore
- process、timeline
- matrix、funnel、roadmap、chart、dashboard
- quote、closing

テーマは `midnight`、`aurora`、`sunset`、`forest`、`minimal`、`ocean`、`berry`、`clay`、`cyber` の9種です。色とフォントは検証済みの範囲で上書きできます。Opusは `design.style`、`density`、`motif` と `variant` を使って、同じ意味レイアウトでも資料固有の視覚表現を選びます。固定PptxGenJSレンダラーはテキスト、図形、組み込みアイコン、テーマ色、編集可能グラフ、グラフ用埋め込みワークブックを生成します。通常の公開入力にファイルパス、URL、画像bytes、JavaScript、任意座標を持たせません。例外となる`media.assetId`は事前登録済みのopaque IDだけで、serverがcaller scopeとSHA-256を検証し、metadataなしPNGを埋め込みます。`visualObjects.assetId`も同じcaller scopeで解決しますが、保存対象は画像ではなく上限付きの意味仕様で、job payloadへimmutable snapshotを保存し、PowerPointネイティブ図形へ展開します。

`media`は最初の実装として`split`だけを提供します。ユーザー提供JPEG/PNGを`pptx_register_uploaded_image_asset`で検証・無害化してから、実asset ID、crop intent、text positionを指定します。画像がない状態を空欄や仮画像で完成扱いせず、native diagramまたは画像なしrecipeへ切り替えます。詳細は[ADR 0016](adr/0016-conversation-scoped-image-assets-and-media-split.md)に記録します。

`nativeDiagram`はtree/flow 3〜12 nodes、cycle 3〜6、concentric 2〜4、network 3〜9、edge最大18を受け、座標をレンダラー側で決めます。既存Processの`loop`、Timeline/Roadmapの`stepped`、Funnelの`pyramid`は、受理されるだけの別名ではなく専用描画分岐を持ちます。補助オブジェクトは`pptx_prepare_visual_objects`へ1回最大8件をまとめ、1ページ最大3件・strong最大1件・会話最大24件です。tool resultはopaque IDと短い説明のJSON textだけを返し、PPTX本体はネイティブ図形です。詳細は[ADR 0017](adr/0017-native-semantic-diagrams-and-visual-objects.md)に記録します。

固定レイアウトはモデルのデザイン判断を置き換えるものではなく、安全に実行できる視覚語彙です。モデルがストーリー、強調対象、構図、視覚モチーフを決め、レンダラーは整列、最小余白、編集可能性、ファイル整合性を保証します。文字量の多い説明では`structuredBrief`が本文を2〜3個の見出し付きセクションへ分け、`scorecard`が評価軸×選択肢を編集可能なPowerPoint表へ変換します。`density=detailed`は単一のフォント倍率ではなく、外周余白、見出し領域、内部間隔、罫線、カード形状、影をまとめて切り替えます。6枚以上で構図が4種類未満、同一構図が3枚連続、文字中心ページが過半数の場合に加え、500文字以上のページでdetailedを使っていない場合や全セクションを強調している場合は `design_warnings` を返します。`VisualSlideSpec.speakerNotes`はvisible canvasとは別に発表者ノートへ保存し、refineで省略した場合は元ページから継承します。詳細は [ADR 0010](adr/0010-readable-information-density.md) と [ADR 0018](adr/0018-speaker-notes.md) に記録します。

`musicScore`は音高、音価、ウクレレの弦・フレット・左手指を意味入力として受け、五線、符頭、符幹、休符、小節線、TAB線、フレット番号、指色マーカーをPowerPointネイティブの図形・線・テキストへ変換します。任意座標や任意描画命令を公開せず、音高とTABの一致、表示密度、対応調弦をドメイン層で検証します。PowerPoint内では個別要素として編集できますが、専用譜面ソフトのような移調・自動組版は行いません。詳細は [ADR 0011](adr/0011-editable-music-score-layout.md) に記録します。

## 自動視覚リフレクション

MCPサーバー指示とツール説明に次のエージェントループを定義します。

1. 生成・編集ジョブを完了させる。
2. 全スライドを1〜4枚ずつMCP画像ブロックとして取得する。
3. Claudeが文字切れ、はみ出し、重なり、文字サイズ、余白、整列、コントラスト、情報階層、密度、バランス、全体一貫性を確認する。さらに、見出しだけで話が追えるか、読む順序が一意か、1ブロック1論点か、強調色が概ね15%以内か、本文が9pt未満に見えないかを確認する。
4. 問題があれば厳密なプレースホルダー資料は `pptx_refine_deck`、白紙資料とブランドVisual Deckは `pptx_refine_visual_slide` へ完全な差し替えページを1枚ずつ渡して再生成する。
5. 最大2巡で収束させ、視覚確認後にダウンロードリンクを提示する。

`pptx_refine_visual_slide` は `jobId=latest` で同じ利用者・会話の最新成功Visual Deckだけを解決します。置換slideが`visualObjects`を省略した場合は元ページの参照をjob snapshotからmaterializeし、異なるIDの明示は拒否します。各Visual DeckジョブはルートID、親ID、修正巡、同巡の修正済みページを永続化します。これにより複数ページの逐次修正を累積しつつ、古い成功ジョブからの分岐、複数ページ一括修正、同一ページの3回目の修正を拒否します。`pptx_refine_visual_deck`は互換用に維持しますが、同じ1ページ制約を適用します。成功後の全体startは拒否し、初回生成が失敗した場合の全体再試行は1回だけ許可します。

ページ数を増やす場合は `pptx_refine_visual_slide` ではなく `pptx_insert_visual_slides` を使います。同ツールは追加する `VisualSlideSpec` だけを受け取り、同じ利用者・会話にある成功済みVisual Deckの仕様へサーバー側で挿入します。`afterSlideNumber` は既存ページを基準とし、省略時は末尾へ追加します。ブランドVisual Deckでは元ジョブのテンプレートファイルと `templateLayoutId` も再利用します。第1段階では完成版の `VisualDeckSpec` をPptxGenJSへ再投入するためファイル生成・Open XML検証・プレビュー生成は全ページ分行いますが、モデルが既存ページを再構築・再送する必要はありません。

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
