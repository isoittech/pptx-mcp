# LibreChat統合

## MCP接続

`integrations/librechat/librechat.fragment.yaml` を既存 `librechat.yaml` へマージします。LibreChat v0.8系では本番MCPに Streamable HTTP を使用し、`mcpSettings.allowedDomains` へコンテナ内のホスト名を追加します。

共有秘密は `apiKey.source: admin` と `authorization_type: bearer` で設定します。LibreChatの起動時インスペクターは管理者提供キーならOAuth検出を省略し、実接続時にだけAuthorizationヘッダーを構成します。`headers.Authorization` を手動設定すると、起動時の未認証プローブをOAuth必須と誤判定するため使用しません。

ユーザーIDと会話IDはLibreChatのヘッダープレースホルダーから渡します。これらをツール引数へ公開しないでください。共有秘密はLibreChatとpptx-mcpの環境変数に同じ値を設定します。

## アップロード連携

現在のアダプターはLibreChatローカルストレージの次の命名規則を読み取り専用マウントから解決します。

```text
/app/uploads/<user-id>/<file-id>__<original-name>.pptx
```

BedrockはPPTX自体をモデル入力へ渡さないため、Claudeからアップロードの`file_id`が見えない場合があります。入力系ツールの`sourceFileId`を省略すると、信頼済みユーザーヘッダーで限定したディレクトリ内の最新PPTXを選択します。`file_id`を明示した場合はそのファイルを優先します。複数のPPTXから特定ファイルを選ぶ必要がある運用では、アップロード直後に処理するか、`file_id`を明示してください。

LibreChat側にも個別ファイル30MBの上限を設定し、リバースプロキシのリクエストボディ上限はmultipartのオーバーヘッドを含めて32MB以上にします。

S3等へ移行する場合は `InputFileResolver` を、LibreChatの認証済み内部ファイル取得APIを呼ぶ実装へ交換します。MCPツールから任意URLを取得する方式にはしません。

## 既定テンプレートと初回案内

既定テンプレートはLibreChatのアップロード領域ではなく、管理者が用意した非公開ディレクトリからMCPコンテナへ読み取り専用でマウントします。`PPTX_MCP_DEFAULT_TEMPLATE_ID` に対応する `<template-id>.pptx` を配置してください。起動時に検証・解析するため、通常の新規作成では利用者による添付も `pptx_analyze` の待機も不要です。

導入環境固有の案内文は `PPTX_MCP_FIRST_ASSISTANT_NOTICE` からMCPサーバー指示へ注入できます。LibreChat側は `serverInstructions: true` を維持してください。これは新規チャット画面の `customWelcome` ではなく、最初のユーザー可視assistant応答に一度だけ案内を含めるようモデルへ指示する方式です。実際の文言は導入側で管理し、OSS設定の既定値は空です。

## 成果物公開

`PPTX_MCP_PUBLIC_BASE_URL` は利用者のブラウザーから到達可能なHTTPS URLです。次の経路だけをリバースプロキシで公開します。

```text
GET /artifacts/{job_id}/{file_name}?token=...
```

`/mcp` はLibreChat内部ネットワークからだけ到達可能にします。署名付きURLをアクセスログへ残さない設定を推奨します。

`integrations/librechat/nginx.conf` をLibreChat側へ配置し、`compose.fragment.yaml` の成果物専用プロキシから読み込みます。このプロキシは `/artifacts/{job_id}/...` のGET/HEADと readiness のみを許可し、`/mcp` を404にします。Dockerの内部ネットワークだけに接続したサービスへはホストポートを公開できないため、MCP本体ではなくこのプロキシだけを通常ネットワークにも接続します。

## Claudeへの運用指示

- 資料編集前に `pptx_analyze` を実行する。
- 会話に`file_id`が提示されていなければ`sourceFileId`を省略し、最新アップロードを使う。
- 通常の新規資料は、全体構成、完成ページ数、テーマ、デザイン、テンプレートを決めてから `pptx_start_visual_deck` を1回呼ぶ。既定テンプレートは`templateSourceFileId=default`、テンプレートなしは`none`とし、start成功後は変更しない。返された`draft_id`へ `pptx_add_visual_slides_to_draft` で連続した1〜4ページずつ追加し、`startSlideNumber`は省略してサーバーに追番させる。`remaining_slide_count=0`になったらfinishを1回だけ呼ぶ。
- 添付した別テンプレートを使う新規資料では、先にそのファイルを解析し、`pptx_start_visual_deck`へ`templateSourceFileId`（不明なら`latest`）と通常`templateLayoutId=auto`を指定する。finishでは選択を上書きしない。startで明示したテーマ色・フォントはテンプレート抽出値より優先される。
- 企業テンプレートの既存プレースホルダー配置へ厳密に流し込むことが明示された場合だけ、解析結果の `layout_id` とプレースホルダー `shape_id` を一字も変更せず `pptx_create_deck` に渡す。動作上必須の`slides`へ完成版の全ページをまとめ、各要素を`layout_id`/`fields`、各フィールドを`text`または`paragraphs`のどちらか一方と`shape_id`（必要時のみ`shape_name`/`placeholder_index`）のsnake_caseで指定する。箇条書きや番号を文字として手入力しない。`sourceFileId`だけで呼ばない。Bedrockから空引数が先行した場合は`input_required`応答を受け、全`slides`付きで直ちに再実行する。
- 新規資料では、内容ごとに意味ベースのレイアウトを選び、`design`とスライド単位の`variant`で構図を指定する。start/add/finishを空引数で呼ばず、addは最大4ページだけを渡す。成功済みVisual Deckがある場合はstart/finishへ戻らず、問題ページの修正または追加ページの挿入を使う。初回生成失敗時の全体再試行は1回までとし、別資料の明示依頼だけ`userRequestedNewWorkflow=true`を使う。
- 成功済みVisual Deckへスライドを追加する場合は`pptx_insert_visual_slides`へ追加分だけを渡す。`jobId=latest`とし、末尾追加なら`afterSlideNumber`を省略する。既存ページを含む完全仕様を再構築せず、作成ツールを呼び直さない。成功状態が直接返れば状態確認ツールを追加せず、全ページの視覚確認へ進む。
- 編集対象が一意でなければ候補を列挙し、利用者の選択まで更新しない。
- ジョブ完了後は `pptx_get_preview_images` で全スライドを1〜4枚ずつ取得し、文字切れ、重なり、可読性、余白、整列、コントラスト、情報階層、密度、バランス、一貫性を評価する。
- 非同期ツールが`job_id`を返したら`pptx_wait_for_job`を使い、`pptx_get_job`を短間隔で反復しない。45秒以内に終わらない場合だけ、同じ`job_id`でもう一度待機する。
- 厳密なプレースホルダー資料に問題があれば`pptx_refine_deck`を使う。白紙資料とブランドVisual Deckは`pptx_refine_visual_slide`へ完全な差し替えページを1枚だけ渡し、ジョブ成功後に次の問題ページを`jobId=latest`で直す。修正は逐次累積し、古いジョブからの分岐、複数ページ一括入力、3巡目はサーバーが拒否する。全ページを最大2巡まで自律的に修正し、上限後に全体再作成へ戻らない。画像を取得していない場合は視覚確認済みと述べない。
- 視覚評価の完了後にPPTXダウンロードリンクを提示する。
- MCPに情報収集を依頼しない。外部情報はLibreChat側で収集し、構造化した内容だけを渡す。

## BedrockでのMCP画像引き渡し

LibreChat v0.8.3-rc1の `@librechat/agents` 3.1.51は、MCP画像を画面用artifactには保存しますが、Bedrockの次回モデル呼び出しには既定で渡しません。導入環境固有のLibreChatパッチは、Anthropic以外の画像対応プロバイダーと同じ `HumanMessage` 変換をBedrockにも適用します。Bedrock変換層は `image_url` のdata URLをConverse APIの画像ブロックへ変換します。

このパッチは通常Dockerfileと `Dockerfile.multi` の双方で `npm ci` 後に適用します。対象依存のコード形状が変わった場合はビルドを失敗させるため、`@librechat/agents` 更新時は上流実装を確認してパッチを削除または更新してください。
