# ADR 0001: Open XML SDKを中核とし商用エンジンを交換可能にする

- 状態: 採用

## 文脈

新規生成だけでなく既存PPTXの未編集要素を保持しながら、SmartArt、編集可能グラフ、埋め込みExcelを更新する必要があります。PptxGenJS と python-pptx は新規生成や通常図形には適しますが、既存PPTXの読み込みとSmartArt編集を要件どおり扱えません。

## 決定

.NET 8 と Open XML SDK を既定エンジンにします。MCPは公式C# SDKのStreamable HTTPを使用します。PPTX操作を `IPresentationEngine` で抽象化し、Aspose.Slidesを使う商用実装を追加できるようにします。LibreOfficeはプレビュー生成にのみ使用します。

## 理由

- Open XML SDK は OOXML パッケージを直接読み書きでき、変更パーツを限定しやすい。
- グラフキャッシュと埋め込みSpreadsheetMLを同じ技術スタックで同期できる。
- SmartArtのDiagram Data Partへ到達できる。
- Microsoftが関与する公式MCP C# SDKと同じランタイムで運用できる。

## 結果

SmartArtノードの整合性維持やPowerPoint固有描画は実装難度が高いため、提供テンプレートで短期スパイクを行います。Open XML実装の工数・互換性が基準を満たさない場合は、Aspose.Slidesの正式ライセンス費用とLinux描画品質を比較して差し替えます。
