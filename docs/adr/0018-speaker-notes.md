# ADR 0018: Visual Deckへ目的主導の発表者ノートを保存する

- Status: Accepted
- Date: 2026-08-16

## Context

スライド上の情報だけでは、発表者が各ページで何を訴えるべきかを再現できない場合がある。visible canvasへ説明を増やすと投影時の可読性を損なう一方、PowerPointの発表者ノートなら、見た目を変えずにページの目的と口頭説明を保持できる。

既存のVisual Deck生成はPptxGenJSが作る空のnotes slideをすべて削除していた。ノートを保持するには、白紙生成のOpen XML正規化、企業テンプレートへの合成、refineの差分契約を同時に更新し、PowerPointが修復を要求しないsingle notes master構造を保証する必要がある。

## Decision

- `VisualSlideSpec`へ任意の`VisualSpeakerNotesSpec`を追加し、`purpose`と`talkScript`を受ける。OSSの既存利用者との互換性を保つため任意とする。
- `purpose`は単一行1〜240文字、`talkScript`は改行を許可する1〜1200文字として検証する。
- rendererは固定見出し「このスライドの狙い」「トークスクリプト」を付け、PptxGenJSの`slide.addNotes`で保存する。visible slide XMLへノート本文を置かない。
- normalizerは空のgenerated notes slideだけを削除する。非空ノートを保持するときは、presentation直下のnotes master relationshipと`p:notesMasterIdLst`を正規化し、notes masterを1件に限定する。
- branded compositionは非空notes slideだけを複製し、テンプレートにnotes masterがあれば再利用する。なければ生成側のmasterを1件だけ取り込み、各notes slideへ共有する。
- `pptx_refine_visual_slide`のreplacementで`speakerNotes`が省略された場合は元ページのノートを継承し、明示値だけを更新する。
- job resultへ`speaker_notes_count`を返し、preview外のノート保持件数を呼出側が確認できるようにする。
- ノートはPPTX受領者が閲覧できるため、内部思考、秘密情報、内部URL、不要な個人情報、未開示の仮定を格納しない。この内容ポリシーは導入環境のAgent／Skillでも強制する。

## Consequences

- 発表者は各ページの訴求点と話す流れをPowerPoint内で確認できる。
- ノート文字列分だけmodel出力、PPTX容量、検証対象が増える。ノートは本文の重複ではなく必要十分な発表原稿に抑える。
- preview画像はノートを表示しない。件数はjob resultで、本文とpackage構造はrenderer／Open XML／E2Eで確認する。
- strict placeholder型の`pptx_create_deck`は今回の対象外とする。

## Validation

- validator testでpurposeの単一行、上限、talkScriptの上限を固定する。
- renderer testでノート本文がnotes XMLにだけ入り、XML特殊文字が安全に保持されることを確認する。
- normalizer testで空ノート削除、非空ノート保持、single notes master、OpenXmlValidatorを確認する。
- branded composition testで企業マスター、ロゴ、フッターを維持したままノートを保持する。
- job refine testで省略時継承と明示更新を確認する。
- ステージングE2Eで全ページの`speaker_notes_count`、ダウンロードPPTXのnotes slide、LibreOffice表示、Open XML検証を確認する。
- 3ページの企業テンプレート付きステージングE2Eで、全slide入力のノート、refine省略時継承、最終`speaker_notes_count=3`を確認した。packageはslide 3件、notes slide 3件、notes master 1件で、固定見出しは各notes slideに1件、visible slideへの漏れは0件だった。ZIP検査とLibreOffice PDF変換にも成功した。
