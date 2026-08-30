# ADR 0020: dom-to-pptxとreact-iconsによるサーバー管理DOMレンダラー

## 状態

採用

## 背景

従来のVisual Deckは、意味仕様を固定PptxGenJSレンダラーへ渡していた。安全性と編集可能性は高い一方、HTML/CSSのレイアウト表現をPowerPointへ変換する工程や、広く保守されているアイコンライブラリを利用できなかった。

一方、モデルにHTML、CSS、JavaScript、SVG、座標を直接生成させると、入力スキーマ、外部リソース、ローカルファイル参照、ブラウザー実行の境界が失われる。既存のネイティブグラフ、表、図解、楽譜をDOM表現へ置き換えると編集性も低下する。

## 決定

- 新規段階生成のレンダラー契約を`visual-v6-dom`とする。保存済み`visual-v4`／`visual-v5` lineageは従来どおり再生成する。
- DOM対応layoutは、サーバーが生成したoffline HTML/CSSから`dom-to-pptx` 2.1.1で変換する。対応範囲と選択基準は[ADR 0021](0021-page-level-visual-composition-and-quality-components.md)および[Visual Component Catalog](../visual-component-catalog.md)を参照する。
- ネイティブグラフ、dashboard、nativeDiagram、musicScoreはPptxGenJSの互換レンダラーを使い、DOM対応ページとページ単位で合成する。これらのネイティブ構造をCSS図形へ退化させず、同じdeckのDOM対応ページまで互換レンダラーへ戻さない。
- モデルが渡すのは従来の`VisualDeckSpec`だけとする。HTML、CSS、JavaScript、任意SVG／XML、座標、URL、path、browser optionを公開tool入力へ追加しない。
- DOMは本文をHTML attributeとtext nodeへエスケープし、外部script、外部stylesheet、外部font、外部画像を参照しない。画像は会話scopeとSHA-256を検証済みのmetadataなしPNG data URIだけを使う。
- Chromiumはコンテナへ固定インストールし、実行時のbrowser自動downloadに依存しない。`dom-to-pptx`のNode exporterは`--no-sandbox`とfile accessを使用するため、非root、read-only、全capability削除、内部network、サーバー生成HTML限定の境界を同時に維持する。
- Cardsの`icon`意味IDは`react-icons` 5.7.0の`react-icons/lu`（Lucide）allowlistへサーバー側で解決する。モデルにReact code、package名、SVG本文を渡させない。DOM経路では同じSVGを使い、互換レンダラーでも`visual-v6-dom` lineageだけはLucide SVGを使う。
- アイコンはベクターSVG画像として出力する。PowerPointの「図形に変換」で分解できるが、変換前から個別DrawingML図形であるとは表現しない。本文、カード、工程等は編集可能なPowerPointテキスト・図形として出力する。
- 企業テンプレート使用時はDOM slide rootを透明にし、従来のOpen XML合成で選択済み白紙レイアウトへ接続する。マスター、ロゴ、フッター、ページ番号を生成側で描き直さない。
- 読ませる本文はPowerPoint内部値で14pt以上にする。dom-to-pptxの換算に合わせてDOM本文は24px以上とし、箇条書きは`ul`/`li`からDrawingMLの`buChar`へ変換する。nativeDiagramのノード・関係ラベルも14pt以上とし、収まらない場合は文字縮小ではなく内容整理またはページ分割を行う。
- 生成直後とテンプレート合成後にOpen XML検証を行い、完成PPTXはLibreOfficeでPDFへ変換して全ページをPNG化する。LibreChatのパワポ職人はOpus 5で全image blockを確認し、問題ページだけ最大2巡修正する。LibreOffice previewをMicrosoft PowerPointのpixel-perfect保証とは扱わない。

## 依存関係と脆弱性境界

PptxGenJS 4.0.1は配布runtimeから呼び出していない`image-size`を推移依存として宣言している。この依存はparserを一切含まず、CommonJS／ESMのどちらから呼ばれても即時失敗するローカルの`@pptx-mcp/image-size-disabled`互換shimへ固定する。これによりICNS／JXL／HEIF parserをruntime imageから除外し、`npm audit --omit=dev`を警告0件にする。PptxGenJSの将来版がこのAPIを実際に使い始めた場合は、黙って解析するのではなくレンダラーを失敗させ、回帰テストで検出する。

公開画像toolは引き続きSharpでmetadataなしPNGへ再エンコードし、rendererはPNG signature、寸法、SHA-256、総bytesを再検証する。URL、path、原upload、ICNS、JXL、HEIFをDOM／PptxGenJSのどちらにも渡さない多層防御を維持する。

## 結果

- 一般的な業務レイアウトはHTML/CSSで保守でき、DOMから編集可能なPPTXを生成できる。
- アイコンの見た目と語彙をLucideへ統一できる。
- ネイティブ構造が重要なページは既存品質を維持できる。
- Chromiumを含むためruntime imageは増える。browser起動とDOM変換の時間も追加されるため、段階生成とジョブtimeoutを維持して監視する。

## 参考

- [dom-to-pptx](https://github.com/atharva9167j/dom-to-pptx)
- [react-icons](https://github.com/react-icons/react-icons)
- [Lucide license](https://github.com/lucide-icons/lucide/blob/main/LICENSE)
