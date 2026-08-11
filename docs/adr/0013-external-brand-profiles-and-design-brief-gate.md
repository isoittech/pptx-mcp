# ADR 0013: 外部Brand ProfileとDesign Briefゲートを導入する

## 状態

採用

## 背景

テンプレートPPTXのマスター、テーマ、数枚のレイアウトだけでは、用途ごとの構図、情報密度、文章トーン、素材の選び方、同一ブランド内の複数方向を十分に表現できない。一方、企業名、ロゴ、実テンプレート、社内URL、承認済み素材の場所をOSSへ含めたり、モデルから任意パスやURLを渡したりしてはならない。

また、Visual Deckはstart時にテンプレート、テーマ、デザインを固定する。生成後にアートディレクションをやり直すと、全体再生成ループ、処理時間、トークン消費、ブランド逸脱が増える。生成前に短いDesign Briefとスライド単位のAsset Planを確定し、startとの順序をサーバー側でも検証する必要がある。

## 決定

- Brand Profileはコンテナ外の読み取り専用ディレクトリから読み込む。各bundleは`<BrandProfilesRoot>/<profile-id>/brand-profile.json`とし、IDは英数字、ハイフン、アンダースコアだけを許可する。ツール入力へファイルパスを公開しない。
- manifestはschema version、profile ID/version、`default`または`none`のテンプレート選択、色の役割、見出し・本文フォント、文章・視覚ルール、禁止・要確認ルール、承認済み素材collection ID、style direction、layout recipe、完成sample要約を持つ。URL、パス、未知フィールド、シンボリックリンクを拒否する。
- 色はprimary、secondary、accent、background、surface、text、muted text、positive、warning、critical、1〜8色のdata seriesを必須とする。textとbackground/surfaceは4.5:1以上、muted textとbackground/surfaceは3.0:1以上をload時に検証する。見出し・本文フォントは別roleとしてレンダラーへ渡す。
- manifestの生バイト列からSHA-256を計算し、宣言versionとcontent hashをcatalogへ返す。catalogはプロセス起動時の検証済みsnapshotを保持し、実行中に外部ファイルが変わっても既存profileの意味を変更しない。変更反映にはプロセス再起動と新しいDesign Brief検証が必要である。
- 公開ツールはschema tokenを抑えるため`pptx_get_design_catalog`へ集約する。無引数ではprofile summaryだけを返す。detail、recipe、sample要約を返す2回目の呼出しでは一覧から選んだ`profileId`を必須とし、`purpose`、`density`、`styleDirectionId`はそのprofile内だけを絞る。profile未指定の絞り込みで全profileのdetailを展開しない。
- `pptx_validate_design_brief`はaudience、purpose、delivery mode、tone、density、immutable profile参照、style direction、visual strategy、source policy、完成ページ数、確定・推定を区別したassumptions、未解決質問、全ページのAsset Planを検証する。`questions_for_user`または`needsConfirmation`が残るbriefは確定しない。
- Asset Planは全ページに1件ずつ必要とし、visual purpose、preferred medium、`nativeDraw` / `userUpload` / `approvedLibrary` / `none`、fallback、license status、crop intent等を宣言する。素材なしは`preferred_medium=none`、`acquisition=none`、`fallback=none`、`status=omitted`、`license_status=notRequired`を正規形とする。第1段階は画像挿入を実装しないため、`userUpload`または`approvedLibrary`を予定する項目だけが`status=fallbackSelected`と`nativeDraw`または`noAssetLayout`のfallbackを選び、外部素材を必須としない別recipeを使わなければならない。URL、パス、画像バイナリ、任意asset IDは受け取らない。
- 検証成功時はランダムなopaque `brief_id`を発行し、利用者scope、会話scope、profile version/hashへ束縛して60分で失効させる。`pptx_start_visual_deck`は`brief_id`指定時にページ数とテンプレート選択を照合し、theme/designの上書きを拒否する。theme/designはprofileのstyle directionとcolor/font rolesから決定する。
- Design Briefがあるdraftでは、各追加slideの`recipeId`をAsset Planと照合し、recipeのsemantic kind、density、実装済みvariantとの一致を強制する。座標、コード、任意レイアウト定義は公開しない。
- 完成deckのjob payloadには、profile version/hash、style direction、ページごとのrecipe契約に加え、秘密を含まないDesign Brief監査snapshotを保存する。snapshotはsource policy、確定・推定assumption、ページごとのvisual purpose、medium、acquisition、fallback、status、license status、検証済みopaque参照、crop指定だけを持ち、期限付き`brief_id`自体は保存しない。refine後もsnapshotとrecipe契約を保持し、第1段階のprofile-bound insertは新規Asset Planを検証できないため拒否する。
- `RequireDesignBrief`はOSS既定で`false`とし、既存クライアントを壊さない。導入環境が`true`にした場合だけstart前のbriefを必須化し、profile rootの欠落または空catalogで起動をfail-closedにする。

## 結果

テンプレート1ファイルだけでなく、ブランドの役割色、文章・視覚ルール、用途・密度別recipe、sample要約を再利用できる。profileを外部設定に保つため、OSSへ会社固有値を持ち込まずに導入環境ごとの語彙を追加できる。Design Briefは単なるプロンプトではなく、startより前の順序、profileの版、素材fallback、スライドrecipeをサーバー側で再現可能に固定する。

代わりに、導入環境はmanifestの作成、version管理、sample説明、recipeとレンダラー実装の整合を保守する必要がある。全ページAsset Planはツール入力を増やすため、catalogはprofile選択後の絞り込み式にし、brief応答は要約だけにする。期限付きbriefと生成途中のdraftは現段階ではメモリ内状態であり、プロセス再起動後は再検証が必要である。完成jobには最小監査snapshotが残るが、元のbrief全文や一時IDは残さない。

第1段階ではsampleの画像表示、承認済みライブラリからの取得、Web画像の権利確認、画像挿入を行わない。これらを追加する場合も、外部取得は導入環境のオーケストレーション側、`pptx-mcp`は権限scope付きopaque file IDの解決と安全な配置だけを担当し、任意URLダウンロードを追加しない。layout recipe、DataTable、実行可能なstyle roleの詳細は [ADR 0014](0014-executable-style-recipes-and-data-tables.md) を参照する。

## 検証

- 外部bundleの厳格JSON、ID、version/hash、URL/path拒否、process内snapshot不変性をテストする。
- 色roleとcontrast、style direction、recipe、sample参照の整合をテストする。
- brief IDの利用者・会話分離、期限切れ、profile hash差異、未解決質問、素材fallback、ページ数・テンプレート・theme/design固定をテストする。
- draft追加時のrecipe ID、semantic kind、density、variant不一致をテストする。
- job payloadのJSON roundtrip後もrecipe契約と監査snapshotが残り、refine時のrecipe逸脱を拒否することをテストする。
- `RequireDesignBrief=false`で従来startが通り、`true`でbriefなしstartが拒否されることをテストする。
