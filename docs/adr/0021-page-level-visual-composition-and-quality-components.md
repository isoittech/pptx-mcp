# ADR 0021: ページ単位の描画合成と品質部品

## 状態

一部置換（新規生成はADR 0022、保存済み`visual-v6-dom`互換と部品語彙は本ADR）

## 背景

ADR 0020では、DOM非対応ページが1枚でも含まれる場合、deck全体をPptxGenJSへ切り替えていた。この方式ではnative chart等の編集性を守れる一方、他の本文ページまで固定実装へ戻り、DOMで表現を改善する余地が失われる。また、実案件で頻出する適用範囲表、変換根拠、成果物一覧、工程表を、既存の汎用カードだけで表すと情報構造と見た目が薄くなる。

## 決定

以下のページ合成と固定DOM描画は`visual-v6-dom`の互換契約として維持する。`visual-v7-author-html`では意味componentを計画・監査用の部品語彙として使い、実際の構図はモデル設計HTML/CSSを正本とする。詳細は[ADR 0022](0022-model-authored-html-slides.md)を参照する。

- `visual-v6-dom`では、DOM対応ページとネイティブ要素ページを1枚ずつ描画し、Open XMLで順番どおりに合成する。
- agenda、bullets、scorecard、dataTableに加え、NativeDiagramのtree/flowもDOM対応へ移す。native chart、dashboard、NativeDiagramのcycle/concentric/network、musicScoreだけを互換描画へ残す。
- 合成時は共通の白紙layoutと単一notes masterへ正規化し、ページ番号には元deckの絶対位置と総ページ数を使う。
- CoverageMap、TransformationEvidence、ArtifactShowcase、GanttScheduleを意味データ契約として追加する。座標、HTML、CSS、任意SVG/XML、URL、pathは公開入力に追加しない。
- DOM契約は影を使用せず、本文を原則14pt以上で設計する。過密なCoverageMapとGanttScheduleはサーバー検証で拒否し、ページ分割を促す。
- Lucideのallowlistを人物、組織、工程、表、日程、成果物、研修等へ拡張する。未知IDは引き続き拒否する。
- 既定テンプレートだけに、導入環境が設定するcover/body見本スライド番号を適用できる。ユーザー指定テンプレートには適用しない。
- 導入環境が本文見出し規則を有効にした場合、既定テンプレートの本文タイトルは30pt Accent 2、主張は別オブジェクトの16pt Accent 2実箇条書きにする。ユーザー指定テンプレートには適用しない。
- DOM移行の検収環境では互換描画を禁止できる。生成結果はページごとのrendererとDOM／fallback件数を返し、fallbackが1ページでもあれば検収失敗とする。
- 詳細な選択基準は[Visual Component Catalog](../visual-component-catalog.md)を単一の参照先とする。

## 結果

- 本文レイアウトの自由度をDOM側へ寄せつつ、グラフ等のネイティブ編集性を失わない。
- 特定の「放置リスク」等へ固定せず、同じ部品を内容に応じて組み替えられる。
- 中間PPTXの生成と合成が増えるため、実行時間と一時ディスク使用量は増える。ジョブtimeout、Open XML検証、保持期限を維持する。
- テンプレートの見本スライドが削除・並べ替えされた場合は、起動・生成時にfail-closedで検出する。
- 内容が衝突していなくても、独立オブジェクト同士が接触・ほぼ接触していれば視覚品質不良として修正する。
