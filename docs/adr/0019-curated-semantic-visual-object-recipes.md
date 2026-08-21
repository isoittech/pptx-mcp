# ADR 0019: 意味ベースの複合Visual Objectレシピ

- Status: Accepted
- Date: 2026-08-20

## 背景

ADR 0017で矢印、枠、吹き出し等を編集可能なPowerPointネイティブ図形として生成できるようにしたが、単体図形のstyle差は線幅、透明度、破線が中心だった。そのため能力は存在してもAgentが標準カードだけで済ませやすく、使った場合も大きな全面枠や吹き出しがプレースホルダーのように見えることがあった。

任意SVG、任意座標、自由曲線をLLMへ公開すると、ブランド逸脱、図形数増加、Open XML不整合、位置修正loopを再導入する。必要なのは自由描画APIではなく、内容上の意味と安全なanchorを固定した少数の構成語彙である。

## 決定

1. 既存`VisualObjectBrief`へ後方互換な`recipe=auto`を追加する。既存jobと入力は単一図形の従来描画を維持する。
2. `directionalCue`、`growthPath`、`focusCorners`、`annotationPin`、`sectionRule`、`cycleCue`の6レシピだけを公開する。各レシピはpurpose、archetype、orientation、placement roleの正規tupleをserverで検証する。
3. レシピは線、矢印端、楕円、角丸ラベル等の複数ネイティブshapeへ展開する。SVG、画像、任意path、座標、生色は入力にもtool結果にも追加しない。
4. content間の`directionalCue`は周囲の見出しが両端を説明するためラベルを禁止する。`focusCorners`は全面枠を描かず四隅だけを示す。`annotationPin`は短いlabelと、折れ線Chart上の1始まりcategory／series番号を必須とする。serverは番号が実データに存在することをslide追加前に検証し、rendererは同じ軸範囲からデータ点を求める。
5. Agentはアウトライン確定後に一度だけ全ページを意味分類し、該当objectを一括prepareする。各ページへの最低使用数は設けず、既存layoutで意味が明確なら`none`とする。
6. profileは既存のallowed archetype/style、ページ・デッキ上限、strong focal ruleでレシピ展開も制約する。複合recipeのstrongは`visualPurpose=emphasis`の`focusCorners`だけをprepare時点で受理し、Design Brief確定後の初回配置で初めて失敗する手戻りを防ぐ。会社固有の色・形状・ロゴをOSSへ追加しない。

## 理由

- PptxGenJSとPowerPoint標準図形の既存能力を利用でき、C#側の認可・監査・snapshot契約も維持できる。
- LLMへ数値座標を考えさせず、同じ意味は同じ構成へ再現できる。
- 全面枠や大きな装飾矢印より控えめで、ブランドテンプレートのロゴ、フッター、本文階層を邪魔しにくい。
- 新しいMCP serverやSVG sanitizerを追加せずに、現在の利用頻度が高い説明表現を改善できる。

## トレードオフ

- 任意イラストや複雑な手描き線は作れない。必要なら別のisolated vector compilerを将来評価する。
- `annotationPin`は折れ線Chartのカテゴリ／系列番号へ追従するが、任意XY座標、棒・円・ドーナツグラフ、複数点注釈は受け付けない。ラベルは安全帯へ置くため、リーダー線の経路はPowerPointの自動plot余白に対する近似を含む。
- schemaとserver instructionsは増える。tool contract byte上限を維持し、レシピを無制限に追加しない。
- レシピを選んでも元layoutの情報設計が悪ければ改善しない。レイアウト選択、文章量、視覚検証が先である。

## 検証

- .NETで6正規tuple、全不一致、annotation anchorの範囲と実Chart照合、非焦点strongの早期拒否、既存`auto`互換、tool schema上限を検証する。
- Nodeで6レシピを1つのPPTXへ生成し、`p:pic`を含まず、矢印端、線、楕円、角丸ラベル、循環矢印がOOXMLへ存在することを検証する。
- LibreOfficeで6ページを画像化し、比較カード間の小矢印、グラフ注釈、四隅枠、見出し罫線、循環キューに重なりや過剰強調がないことを確認する。
- PowerPoint互換は既存の生成後OpenXmlValidatorとruntime buildで回帰確認する。
- ステージングの2ページE2Eで、折れ線Chartのcategory 3／series 1へannotationPinを束縛し、preview上で71%のmarkerと注釈始点が重なること、refine後も同じserver-owned assetを継承することを確認する。
