# ADR 0021: ページ単位の描画合成と品質部品

## 状態

採用

## 背景

ADR 0020では、DOM非対応ページが1枚でも含まれる場合、deck全体をPptxGenJSへ切り替えていた。この方式ではnative chart等の編集性を守れる一方、他の本文ページまで固定実装へ戻り、DOMで表現を改善する余地が失われる。また、実案件で頻出する適用範囲表、変換根拠、成果物一覧、工程表を、既存の汎用カードだけで表すと情報構造と見た目が薄くなる。

## 決定

- `visual-v6-dom`では、DOM対応ページとネイティブ要素ページを1枚ずつ描画し、Open XMLで順番どおりに合成する。
- agenda、bullets、scorecard、dataTableもDOM対応へ移し、native chart、dashboard、nativeDiagram、musicScoreだけを互換描画へ残す。
- 合成時は共通の白紙layoutと単一notes masterへ正規化し、ページ番号には元deckの絶対位置と総ページ数を使う。
- CoverageMap、TransformationEvidence、ArtifactShowcase、GanttScheduleを意味データ契約として追加する。座標、HTML、CSS、任意SVG/XML、URL、pathは公開入力に追加しない。
- DOM契約は影を使用せず、本文を原則14pt以上で設計する。過密なCoverageMapとGanttScheduleはサーバー検証で拒否し、ページ分割を促す。
- Lucideのallowlistを人物、組織、工程、表、日程、成果物、研修等へ拡張する。未知IDは引き続き拒否する。
- 既定テンプレートだけに、導入環境が設定するcover/body見本スライド番号を適用できる。ユーザー指定テンプレートには適用しない。
- 詳細な選択基準は[Visual Component Catalog](../visual-component-catalog.md)を単一の参照先とする。

## 結果

- 本文レイアウトの自由度をDOM側へ寄せつつ、グラフ等のネイティブ編集性を失わない。
- 特定の「放置リスク」等へ固定せず、同じ部品を内容に応じて組み替えられる。
- 中間PPTXの生成と合成が増えるため、実行時間と一時ディスク使用量は増える。ジョブtimeout、Open XML検証、保持期限を維持する。
- テンプレートの見本スライドが削除・並べ替えされた場合は、起動・生成時にfail-closedで検出する。
