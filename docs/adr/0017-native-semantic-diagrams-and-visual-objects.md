# ADR 0017: 意味ベースのネイティブ図解とVisual Object asset

- Status: Accepted
- Date: 2026-08-16

## 背景

Visual Deckは企業テーマと既存22 layoutを安定して再現できる一方、工程・関係・循環・範囲を同じカード列へ寄せやすかった。LLMへ「センスよく」「矢印を追加」とだけ指示しても、座標調整、過剰装飾、同じtoolの反復を招き、再帰上限やOpen XML不整合の再発要因になる。

参考にしたPowerPoint解説の本文・リンク先14記事・掲載スライド画像では、見栄えは自由曲線の多用ではなく、意味に合う標準図形、揃った寸法と間隔、短い見出し、少数の接続線、一点だけの強調によって成立していた。矢印は方向・成長・時間・循環を、枠と吹き出しは特定の判断・データを示すときだけ使われていた。この表現原則を抽象化し、画像自体は複製しない。

## 決定

1. `VisualSlideKind.NativeDiagram`を追加し、tree、flow、cycle、concentric、networkを意味入力からPowerPointネイティブ図形へ変換する。座標は公開せず、nodeはkind別2〜12件、edgeは最大18件、tree/flowはacyclicを強制する。
2. 既存layoutへ、見た目が実際に変わるvariantだけを追加する。Process=`loop`、Timeline/Roadmap=`stepped`、Funnel=`pyramid`を3〜6 itemsで受理する。公開variantが`auto`へ黙って退化する実装は禁止する。
3. 補助図形は別tool `pptx_prepare_visual_objects`で1〜8件をまとめて準備する。入力はvisual purpose、archetype、style、emphasis、orientation、placement role、palette role、任意の短いlabelだけとし、座標、生色、SVG/XML、URL、path、コードを受け取らない。
4. Visual Object assetはuser/conversationへ束縛し、1時間で失効する。1ページ最大3件、strong最大1件、会話最大24件とする。返したIDをDesign Briefの同じslideの`visual_object_asset_ids`へ固定する。add時にslide `visualObjects`を省略した場合はサーバーが検証済み計画からmaterializeする。refine時も省略すればjob snapshotから元ページのIDをmaterializeし、明示された異なるIDは拒否する。
5. job payloadへID、意味仕様、SHA-256 fingerprintのimmutable snapshotを保存する。PPTX生成と後続refineはこのsnapshotから再現し、期限付きrepository stateへ依存しない。
6. Brand Profileは`visual_object_policy`でallowed archetype/style、default style、page/deck上限、strong focal ruleを外部設定できる。会社固有値はOSSへ含めない。
7. tool結果はopaque IDと短い説明のJSON textだけを返す。SVGをMCP ImageContentとして返すと、現行Bedrock／Anthropic経路が次stepで未対応画像として処理しprovider errorになるため公開しない。PPTX本体は個別編集できるネイティブshape/line/textとする。任意SVG生成はこのMCPへ入れず、将来必要になった場合にisolated renderer/serviceとして別途評価する。

## 理由

- Native-firstはPowerPoint内編集、企業色role、Open XML検証、アクセシビリティを維持しやすい。
- batch prepareはobjectごとのtool往復を避け、LLMの反復回数とtokenを抑える。
- 意味入力と上限を先に固定すれば、座標の微修正ループ、巨大schema、図形数爆発を防げる。
- 23 layoutを機械的に増やすより、既存layoutの実variantと不足していた関係図を足す方が保守コストに対する効果が高い。

## トレードオフ

- 自由なSVGより表現力は狭い。複雑なイラスト、ブランド固有ornament、任意pathは生成できない。
- placement roleは自動配置のため、任意のデータ点をpixel単位で指す用途には向かない。必要ならpage preview後に1ページだけrefineする。
- placement roleは汎用座標へ空図形を置く契約ではない。吹き出しはlabelを図形内へ入れ、focus frameは既存の主要領域へ沿わせ、top-to-bottom NativeDiagramの括弧は下段node上の安全域で関係を束ねる。意味layout別anchorを定義できない組合せは、将来のschema拡張なしに派手なプレースホルダへ退化させない。
- in-memory assetはserver再起動で失われるため、Design Brief validation前の未使用IDは再準備が必要になる。一度draftへ追加されたobjectはjob snapshotに残る。
- tool結果には画像previewがない。archetype、style、emphasisは短い構造化説明で確認し、最終外観はPPTX全ページpreviewで検証する。

## 検証

- .NETでbatch/page/conversation上限、semantic mismatch、cross-user/conversation、expiry、Design Brief binding、snapshot persistence、diagram node/edge/cycleを試験する。
- Nodeで全diagram kind、全新variant、prepared objectを1 deckへ描画し、`p:pic`を含まずnative shapeとarrow endpointを持つことを検証する。
- LibreOfficeで代表deckをPNG化し、枠と矢印が控えめで、重なり・切れ・過剰強調がないことを確認する。
- 導入環境のchat UIで1回のbatch prepare、Design Brief、start/add/finish、全ページpreview、必要ページのみ最大2巡refineをE2E確認する。
- 7ページE2Eで見つかった括弧と下段nodeの重なりを、node間隔、connector終端、layout別anchorへ修正した。修正後の3ページE2Eは142秒で、catalog 2回、batch prepare 1回、add 1回、preview 2回、`visualObjects`省略refine 1回をエラーなく完走した。
- 最終PPTXは3ページ、native shape 54、slide picture 0、外部relationship 0で、ZIP整合とLibreOffice変換に合格した。7ページE2Eの569秒は性能課題として残し、再帰上限の引上げで隠さない。
