# ADR 0014: デザイン方向を実効トークンにし、汎用表とページ密度を追加する

- 状態: 採用

## 文脈

Visual Deckは20種類の意味レイアウト、5つの`design.style`、6つの`variant`を公開していた。しかし、完成スライドとrendererを照合すると、`style`の差は主にカードの角と影に限られ、`grid`と`cascade`は描画へ影響しなかった。参考資料155ページでは、用途に応じた低・中・高密度が同じ資料内に共存し、編集可能な表が27個使われていたが、現行はdeck-globalなdensityと評価表専用の`Scorecard`しか持たなかった。

受理したデザイン指定を黙って無視する状態は、Agentの選択、Brand Profileのレシピ、視覚回帰テストの信頼性を損なう。汎用表を`Scorecard`へ偽装すると、評価軸×選択肢という意味契約も崩れる。

## 決定

- 5つの`design.style`を固定rendererの実効style profileへ対応させる。
  - `executive`: 大きい見出し、抑制した影、明確な結論線。
  - `editorial`: 直線的な面、影なし、細いaccent rule。
  - `bold`: 強いside band、太いaccent rule、強めのshadow。
  - `technical`: 矩形、薄いgrid、細いborder、高密度に適した面。
  - `playful`: 丸み、2色のmarker、軽いshadow。
- style profileは既知の色役割、形状、線、影、文字倍率だけを持つ。任意座標、任意コード、URL、pathを受け付けない。
- 描画されない`grid`と`cascade`を新規`visual-v5`入力の公開`variant`から削除する。保存済み`visual-v4` lineageは再描画・refine時の互換性のため受理する。
- `split`は4件以上かつtakeawayなしの`Bullets`、`spotlight`は正確に3件の`Metrics`または3〜4件の`Cards`、`editorial`は正確に3 sectionsの`StructuredBrief`だけで受理し、kind・件数・状態の不一致を検証エラーにする。Brand Profileの`Metrics` spotlight recipeは`item_count: 3`も必須とする。
- `VisualSlideSpec.density`を追加し、未指定時だけdeckの`design.density`を継承する。rendererの余白、間隔、表、文字倍率はページごとの実効densityを使う。
- `DataTable`を21番目の意味レイアウトとして追加する。
  - 2〜6列、1〜10行のPowerPointネイティブ表とする。
  - 列見出し、alignment、相対幅、行セル、意味色、限定的なemphasisを扱う。
  - 全行のcell数を列数と一致させる。
  - column、row、cellのnullを検証エラーとし、headerとcellの明示改行を禁止する。
  - densityごとに列数、行数、総文字数を制限し、9pt未満への一括縮小や自動改ページを行わない。
- `recipeId`を任意の不透明識別子としてslideへ保持する。rendererへ任意処理を与えず、active Brand Profileとの整合検証と監査にだけ使う。
- renderer contractをv5へ更新し、DataTableの生成テスト、代表style／densityのOOXML契約テスト、LibreOfficeによる手動視覚確認を追加する。
- `visual-v5`ではsurface上の全contentにon-surface foregroundを使い、primary、secondary、accent上の前景も実contrastから選ぶ。既存jobの再描画結果を変えないため、`visual-v4`の固定foreground、semantic tone、card面の色契約は維持してXML回帰テストで固定する。

この決定は[ADR 0003](0003-expressive-visual-language.md)のレイアウト数とvariant一覧を更新し、[ADR 0010](0010-readable-information-density.md)のdeck-global densityをページ単位へ拡張する。

## 理由

- 名前だけのデザイン方向ではなく、同じtemplate内で構造的に異なる資料を決定的に生成できる。
- 情報量の少ない区切りと高密度の詳細表を同じdeckで扱える。
- 汎用の状態表、要件表、スケジュール表を、比較評価の`Scorecard`と区別して編集可能にできる。
- 無効な指定を早期に拒否し、Agentが存在しない表現差へtokenと修正巡を使うことを防げる。
- 外部Brand Profileは、rendererが実際に保証するtokenとvariantだけをrecipeとして公開できる。

## 不採用案

- `grid`と`cascade`を名前だけ残す: 受理済み入力を黙って無視するため不採用とする。
- 表を画像として貼る: PowerPoint上の編集可能性とアクセシビリティを失うため不採用とする。
- 任意の行列数を自動縮小する: 文字切れ、極小文字、PowerPoint修復リスクを事前に制御できないため不採用とする。
- 23用途をそれぞれ新しいenumにする: Q&A、まとめ、次アクション等は既存kindのrecipeで表現でき、schemaとtest matrixだけを増やすため不採用とする。

## 結果

新規`visual-v5`入力で`grid`、`cascade`、またはkind・件数・状態に合わないvariantを使った場合は明示的な検証エラーになる。保存済み`visual-v4` lineageの`grid`と`cascade`は互換性のため受理するが、既存ページの見た目を変える新しい効果は付与しない。

style profileとDataTableの追加によりrendererの視覚回帰対象は増える。現段階は高リスクの表、surface contrast、色role、v4/v5互換性をOOXML契約テストで固定し、代表資料をLibreOfficeで目視確認する。全style×全kind×全densityの直積を毎回生成せず、将来のgolden setは代表的な低・中・高密度へ絞る。写真や複数画像は安全なasset pipelineが完成するまで追加しない。
