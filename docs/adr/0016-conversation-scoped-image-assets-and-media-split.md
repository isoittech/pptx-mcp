# ADR 0016: 会話scope付き画像assetとMedia splitを導入する

## 状態

採用

## 背景

Asset Planは写真・イラストの必要性、取得元、権利、crop、text safe areaを表現できたが、`userUpload`も常に画像なしfallbackへ落としていた。このため画像が必要な構図を実行できず、空欄や仮画像を完成扱いしないという方針は守れても、利用者が承認済み画像を添付した場合に活用できなかった。

`pptx-mcp`は外部networkへ接続せず、URLや任意pathを入力にしない。LibreChat upload領域のファイルをそのままPPTXへ渡すと、形式偽装、巨大画像、metadata、symlink、別利用者・別会話への流用、外部relationship、保持期間の不整合が残る。

## 決定

- 第1縦切りは`userUpload`のJPEG/PNGと、単一画像を使う`Media/split`に限定する。approved library、Web取得、画像生成、複数画像galleryは別段階とする。
- 公開tool `pptx_register_uploaded_image_asset`はLibreChatのopaque `file_id`または`latest`、代替テキスト、任意のopaque attribution IDだけを受け取る。URL、path、ファイル名、画像bytes、自由形式の出典URLを受け取らない。
- LibreChatのmessage attachment画像は通常の文書uploadとは別の`images/<user-id>/`へ保存される。導入環境は画像rootと文書upload rootを別々に読み取り専用mountし、resolverは両方の同一caller user ID配下だけを探索する。
- uploadはtrusted callerのuser ID配下だけを参照し、regular file、symlinkなし、opaque ID、拡張子、JPEG/PNG signature、12MiB上限を検証する。
- 原本はserver-owned Node 20 + Sharpでdecodeし、単一frame、20Mpx以下、長辺2560px以下へ制限する。orientationを反映し、sRGB・metadataなしのPNGへ変換する。PPTXへ原本を直接渡さない。
- 無害化済みPNG、幅、高さ、SHA-256、alt text、acquisition、license status、opaque attribution、user scope hash、conversation scope hash、作成・失効時刻をserver-owned storageへ保存する。asset IDはlowercase 32 hexとし、保持期間はjob既定と同じ7日、期限後は定期削除する。
- `AssetPlanItem.asset_id`は`acquisition=userUpload`かつ`status=ready`のときだけ使う。`license_status=userProvided`、`fallback=none`、photo/illustration、crop intent、`text_safe_area=left|right`、required asset roleを持つrecipeを必須にする。`approvedLibrary`は引き続きfallback-onlyである。
- `VisualSlideKind.Media`と`VisualMediaSpec`を追加する。最初の実装variantは`split`だけとし、実asset ID、crop intent、text positionを必須にする。空欄、仮画像、URL、pathはMedia payloadとして表現できない。
- draft追加時とjob submit/refine時にuser/conversation ownershipと期限を再検証する。Brand Profile-bound deckではAsset Planのasset ID、crop、text safe area、attributionと一致させ、refineでも固定する。
- rendererへはserver-owned asset metadataとrootを別経路で渡し、SHA-256を再検証してdata URIとしてPPTX packageへ埋め込む。外部relationshipを作らない。alt textはPowerPointのpicture descriptionへ設定する。
- PptxGenJS配下の`image-size`にはICNS／JXL／HEIF parserのDoS advisoryが残る。rendererはSHA-256に加えてPNG signature、IHDR、宣言寸法を再検証し、対象形式をparserへ到達させない。依存監査警告は到達不能化だけで解消済みとせず、上流更新を追跡する。
- cropは`contain`、`cover`、`focalCenter`、`focalLeft`、`focalRight`を実装し、text positionはleft/rightとする。画像が利用不能ならMediaを完成扱いせず、native diagramまたは画像なしrecipeへ戻す。
- deckあたりassetは20件、無害化済みbytes合計64MiBまでとする。登録時の画像形式検査と、renderer時のintegrity検査を別々に維持する。

## セキュリティ・権利境界

- `userProvided`は「利用者が提供した」という監査状態であり、著作権や肖像権の自動承認ではない。Agentは利用権を確認し、人物・商標・製品・機密画像は必要なら利用者へ承認を求める。
- asset IDの秘密性を認可根拠にせず、caller user/conversation scopeを毎回照合する。
- serverは外部networkへ接続せず、画像URLをdereferenceしない。Web素材は将来も導入環境側brokerが取得・審査・保存したopaque IDだけを渡す。
- Sharpのdecodeは技術的無害化であり、画像内容の真正性、誤情報、機密性、権利確認の代替ではない。
- PPTX job payloadと監査snapshotにはopaque asset IDと出典record IDが残る。原upload path、内部URL、署名付きURL、元metadataは残さない。

## 却下した案

- `pptx-mcp`がURLを受けてdownloadする: network・SSRF・権利境界を壊すため却下する。
- upload原本をそのままPptxGenJSへ渡す: 形式偽装、metadata、parser負荷をrenderer境界へ持ち込むため却下する。
- 「ここに画像」shapeをMediaとして完成させる: 素材不足を隠し、受入条件に反するため却下する。
- 最初からgallery、人物切り抜き、背景除去、生成画像まで扱う: schema、rights、crop、視覚回帰が同時に広がるため分離する。
- asset IDだけでcallerを認可する: 会話共有やID漏えい時に横断利用できるため却下する。

## 残存課題

- approved libraryはcollection IDの計画だけで、asset解決・ACL・版管理は未実装である。
- Web検索・取得、原典・license確認、画像生成は未実装である。
- `Media`はsplit 1枚画像だけで、gallery、背景全面、evidence strip、人物／製品専用cropは未実装である。
- user uploadの権利確認は利用者申告とAgent運用に依存し、承認workflowやrights ledger UIはない。
- asset retentionとPPTX/job retentionは同日数だが、既存PPTXへ埋め込まれた画像はPPTX保持期間に従う。元asset削除後も生成済みPPTXから消えるわけではない。

## 検証

- JPEG/PNG signature、拡張子不一致、symlink、size/pixel/dimension、decode失敗、metadata除去をtestする。
- user/conversation/expiry、asset ID tamper、SHA-256不一致、合計size上限をtestする。
- `Media/split`のOpen XMLへembedded image relationship、alt text、crop rectがあり、external relationshipがないことをtestする。
- placeholder-only Media、Asset Plan不一致、refineでのasset/crop/text position変更を拒否する。
- runtime containerでLibreChatの実message attachment rootからregister→Design Brief→draft→finish→LibreOffice previewをE2Eし、PowerPoint修復警告相当のOpen XML検証を通す。
