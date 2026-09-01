# ADR 0023: モデル生成HTML/CSSの検証失敗をページ単位で復旧する

## 状態

採用

## 背景

`visual-v7-author-html`では、会話モデルが複数回に分けて完成HTML/CSSをdraftへ追加し、最後にPPTXへ変換する。従来は既定テンプレートのprotected roleに必要なfont-size指定の一部をNode rendererだけが検査していた。このため全ページ受理後のfinishで初めて誤りが分かり、失敗jobは成功済みjob専用のページ修正toolでも直せなかった。会話モデルが新しいDesign Briefとdraftを開始すると、既に作った正常ページまで再生成され、tool回数とcontextを浪費する。

## 決定

- 既定表紙と本文のprotected role font-size契約は、`pptx_add_visual_slides_to_draft`がbatchを受理する前に検証する。不正batchは原子的に拒否し、含まれるページを1枚も保存しない。
- Node rendererはmodel-authored HTML/CSSの検証失敗を`visual_authored_html_invalid`へ分類し、ページ別変換ではエラーへ1始まりページ番号を付ける。
- `visual_authored_html_invalid`と、同じ原因だと確実に識別できる旧`visual_renderer_failed`は、失敗jobの保存済みpayloadをcheckpointとして扱う。`pptx_refine_visual_slide`はそのjobから指摘ページ1枚だけを差し替え、全体を再レンダリングできる。
- 回復可能なHTML/CSS失敗後の暗黙的な`pptx_start_visual_deck`は拒否する。status toolは失敗job IDとページ単位修正の指示を返し、新しいDesign Brief、別draft、正常ページの再送へ誘導しない。
- Open XML整合性、Chromium起動、タイムアウト等の内部障害を、モデルがHTML/CSS変更で直せる失敗とは扱わない。

## 結果

利用者が確認済みの構成と正常ページを保持したまま、誤った1ページだけを修正できる。入力契約違反は早く返り、finishまで進んでから全体が失敗する頻度も下がる。一方、過去のgeneric errorを復旧対象に含める判定は、error messageがmodel-authored HTML/CSSまたは既定テンプレート契約を明示する場合に限定する。
