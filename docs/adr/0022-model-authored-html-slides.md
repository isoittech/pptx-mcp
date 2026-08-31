# ADR 0022: モデル設計HTML/CSSをPowerPointへ変換する

## 状態

採用

## 背景

`visual-v6-dom`はモデルが意味JSONを選び、サーバー側の固定rendererがHTML/CSSを組み立てていた。この方式は安全で予測可能だが、モデルのデザイン判断がlayout名やnode構造に限定されるため、余白、情報階層、図形の組合せ、ページ固有の視覚表現を十分に発揮できなかった。結果として、dom-to-pptxを使用していても従来の固定レンダラーと似た見た目になった。

参考にした[Zenn記事](https://zenn.dev/flux/articles/5446dba24b5100)は、AIがまずスライドをHTML/CSSとして設計し、ブラウザーの計算済みDOMをdom-to-pptxで編集可能なPowerPointへ変換する工程を示している。この役割分担を新規生成へ採用する。

## 決定

- 導入設定`UseModelAuthoredHtmlRenderer=true`の新規draftは`visual-v7-author-html`を使う。保存済みv4〜v6 lineageの再生成契約は変更しない。
- 会話モデルは各ページを1600×900の静的Webページとして設計し、`VisualSlideSpec.authoredHtml.html`へfragment、`authoredHtml.css`へ`.slide`配下のCSSを渡す。HTML/CSSが見た目の正本である。
- `kind`、recipe、density、variantは計画・監査用に保持するが、visible contentを旧renderer用の表・カード・図解payloadへ二重記載しない。保存済み再試行に旧payloadが含まれても、v7では描画せず、そのrenderer固有geometry検査も行わない。
- serverはモデルの構図を固定カード、tree、diagramへ再解釈しない。外側の`.slide`、テーマ変数、企業テンプレート合成、安全性検証、素材解決、Chromium描画、dom-to-pptx変換だけを担当する。
- 意味layout、recipe、density、variant、title、speaker notesは計画、監査、構造検査のため残すが、server側の固定構図を選ぶ指示として使わない。
- 1回のaddでは、各HTML/CSSが完結した連続2〜4ページを渡し、残り1ページだけは単独で渡す。完成HTML/CSSは出力量が大きいため通常は2ページを1batchとし、3〜4ページは短いページだけを完全なtool引数が応答枠へ十分収まる場合に選ぶ。ページ数を増やすための省略は行わない。refineでは問題ページの完成HTML/CSSを丸ごと差し替える。
- 全ページをdom-to-pptxで変換し、PptxGenJS互換rendererへfallbackしない。job resultの`renderer_usage_by_slide`、`dom_rendered_slide_count`、`fallback_rendered_slide_count`を検収に使う。
- 既定企業テンプレートでは表紙見本と本文見本のlayoutへDOM overlayを合成する。利用者指定テンプレートには既定見本番号や既定見出し規則を流用しない。
- NTT DATA既定本文規則を有効にした環境では、`slide.title`／`slide.subtitle`と同じ表示文字列を、正確に1個の`h1|h2[data-pptx-role="body-title"]`と正確に1個の単一項目`ul[data-pptx-role="body-claim"]`へ分ける。モデルは最終表示値の50px／27pxで配置を設計し、serverの最終CSSでタイトル30pt相当、主張16pt相当の実箇条書き、ともにAccent 2を保証する。小さい文字で配置してからserver側だけで拡大する不整合は検証で拒否する。
- 保護roleの文字サイズはrole属性を直接末尾に持つCSS selectorのほか、そのrole要素だけに一意に使われるclass selectorでも検証する。同じclassが通常本文へ再利用されている場合は例外サイズを認めない。HTML上のroleとCSS上の実対象が同じでもclass selectorを使っただけで誤拒否することを避けつつ、12pt相当の出典例外が本文へ波及しないようにする。

## 安全性境界

- HTMLはparse5でparseし、許可tagと許可attributeだけを新しいfragmentへ再構築する。要素数、文字数、CSS量、素材数に上限を設ける。
- script、style tag、JavaScript、event属性、外部URL、data URI、ローカルpath、任意SVG/XML、iframe、object、embed、formを拒否する。
- CSSは全selectorが`.slide`から始まる場合だけ受理し、ページ固有属性へscopeする。`@` rule、`url()`、`expression`、fixed配置、疑似要素content、animation、transition、box/text shadowを拒否する。
- `.slide`は24pxを継承し、モデルが明示する`font-size`はpx単位かつ24px以上だけを受理する。出典、URL、発行日の要素自身へ`data-pptx-role="source-meta"`を付けた場合だけ20pxまで許容する。
- Lucideは`data-pptx-icon`の承認済み意味IDをserver allowlistから解決する。React codeとSVG本文はモデルから受け取らない。
- 画像は事前登録済みopaque IDだけを`data-pptx-asset`で参照し、caller scope、SHA-256、PNG形式、alt textをserverが再検証する。
- Chromiumを実行するコンテナは非root、read-only、全capability削除、内部networkとし、検証済みoffline HTMLだけを開く。

## 品質確認

- dom-to-pptxが未対応または挙動差のある複雑なCSSは使用せず、変換実績のあるHTML要素、grid、flex、absolute配置、単純なborder／fill／typographyを中心にする。
- ChromiumとPowerPoint／LibreOfficeの文字メトリクス差に備え、タイトル、主張、次の独立ブロック間に20px以上の見える余白を置き、文字量の多い領域は高さを使い切らず概ね15%の予備を残す。
- 生成後はLibreOfficeで全ページを画像化し、Opusが見切れだけでなく、情報階層、余白、主張、内容密度、原資料忠実性を確認する。ただし画像だけではpt値や実箇条書き構造を判定できないため、Open XMLの構造検査を必ず併用する。LibreOfficeとMicrosoft PowerPointの文字組み・gradient差は区別し、最終表示の正本はPowerPointとする。
- 独立オブジェクトの接触、本文14pt未満、既定本文タイトル30pt未満、タイトルと主張の結合、fallback発生を検収不合格とする。

## 結果

- モデルがHTML/CSSの表現力を直接使ってページ固有のデザインを作れる。
- server側rendererを増やすことがデザイン改善の前提ではなくなり、固定金型への逆戻りを防げる。
- モデル出力が長くなり、CSS互換性と安全性検証が重要になる。各ページのHTML/CSSを独立して完結させ、最大4ページのbounded batch、fail-closed検証、プレビュー修正を必須とする。
