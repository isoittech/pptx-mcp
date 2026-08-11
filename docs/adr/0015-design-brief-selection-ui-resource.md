# ADR 0015: 条件付きDesign Brief選択カードをlegacy UI Resourceで提供する

## 状態

採用

## 背景

Brand ProfileとDesign Briefゲートは、生成開始前にブランド、用途、素材fallback、ページrecipeを固定できる。一方、方向性の選択が結果へ大きく影響する場合も、従来はAgentが推奨案を文章で説明してそのまま生成するため、利用者は完成後まで別案の存在と判断根拠を確認できなかった。

すべての依頼へ確認画面を挟むと、明確な依頼にも人間の往復、待ち時間、token消費が増える。候補全文を複製すると、最大50ページのAsset Planが案数分だけ入力へ重複する。クライアント値を信用してstyleやbriefを切り替えると、改ざん、replay、別利用者・別会話への流用が生成境界を迂回する。

現行LibreChatホストは`@mcp-ui/client` 5.7.0のinline resourceを扱うが、stableな公式MCP Appsのpredeclared resourceとJSON-RPC action contractではない。この互換範囲を過大に表現せず、現行ホストで安全に段階導入する必要がある。

## 決定

### 条件付きフロー

- 目的・読者・利用場面・方向が明確、safe defaultで十分、または実行可能な案が1件だけの場合は、従来どおり`pptx_validate_design_brief`から`pptx_start_visual_deck`へ進む。カードを儀式として常時表示しない。
- 結果へ大きく影響し、実際のレンダリングが異なる案を2件以上提示できる場合だけ、次の順序にする。
  1. `pptx_prepare_design_brief`
  2. `text/html`のinline UI Resourceを提示し、そのturnを終了
  3. 利用者の固定intentを待つ
  4. `pptx_apply_design_brief_action`
  5. applyが返した`brief_id`だけで`pptx_start_visual_deck`
- 公開actionは`{type:"intent",payload:{intent:"pptx.designBrief.select",params:{choiceSessionId,optionId}}}`に固定する。apply toolは`choiceSessionId`と`optionId`以外を受け取らない。clientからbrief ID、style、action名、nonce、URL、pathを受け取らない。
- 推奨briefと全ページAsset Planは1件だけ渡し、style別案は`style_direction_id`、density、slide順recipe IDだけの差分とする。画像を使わない別構成は、写真計画ページだけのcanonical overrideとする。
- 選択肢は推奨を含め最大3件、最小2件とする。画像を使わない別構成がある場合はstyle別案を最大1件にする。
- 別案はIDや表示名の違いではなく、実効renderer fingerprintで比較する。profileの全color roleがpresetを上書きする現schemaでは`theme_preset`差を実効palette差と数えない。style、motif、色・font role、各ページのsemantic kind、variant、density、item countをrendererと同じcase-insensitive tokenへ正規化する。Brand Profile-bound slideはrecipe densityを強制するため、deck base densityだけの差も別案と数えない。
- 現schemaはprofile全体でcolor roleが共通であり、style direction別paletteは表現できない。方向差はstyle、motif、ページrecipeの構造差を中心とし、方向別paletteは後続schema拡張とする。

### UI Resource

- UIはself-containedな`text/html`とし、外部URL、外部network、任意local path、外部script/font/imageを含めない。
- CSPは`default-src 'none'`、`connect-src 'none'`、`img-src data:`等を指定し、固定inline scriptはSHA-256 hashで許可する。全profile・brief表示文字列をHTML encodeする。
- scriptは固定intentだけをpostし、heightだけを240〜900pxへclampして通知する。widthはhostへ送らない。
- カードは推奨style、目的、読者、delivery mode、density、visual strategy、tone、Asset Plan要約、最大3件の主な前提、source policy、完成sample carouselを短く表示する。
- CTAは一度選ぶと方針が確定し生成へ進むことを明示する。apply後の別案変更は許可しない。
- 第2段階では外部画像をPPTXへ挿入しない。写真計画も検証時に画像不要recipeのfallbackへ確定しているため、「写真あり対写真なし」とは表示せず、「画像を使わない別構成」と表示する。真の写真比較はAsset Broker導入後に扱う。
- toolのmodel-facing textはUI Resource markerの提示、turn終了、pending中のvalidate/start禁止を指示する。option ID対応をplain textへ展開しない。UI非対応hostでは自動選択させずfail-closedにする。
- カードが表示できない、または利用者がカードなしのsafe defaultを明示した場合に限り、引数なしの`pptx_cancel_design_brief_selection`で同じ利用者・会話の未選択pendingを破棄できる。選択済み、start予約中、start済みは取消不可とする。cancel後のdirect validationは`selection_source=agentDefault`になる。

### 選択とstartの状態境界

- `choiceSessionId`と`optionId`はランダムなopaque IDとし、server stateだけで候補へ解決する。sessionはuser scope、conversation scope、有効期限、profile ID/version/content hash、許可optionへ束縛する。
- 同じoptionの二重clickはstart前だけ同じbriefを返す。別optionの再送はreplayとして拒否し、start後の古いカードはalready-startedとして拒否する。
- 1利用者・1会話に未選択pendingは1件だけとする。明示的なreplaceは未選択pendingにだけ許可し、apply後は置換しない。利用者単位とcatalog全体の上限も設ける。
- prepareは同じ会話の未開始briefを無効化し、Pending gateを設定する。Pending中は`RequireDesignBrief=false`でも、既存brief、任意brief ID、briefなしstartをすべて拒否する。
- apply後は選択されたbriefだけstart可能にする。AuthorizeとBeginの間はconversation単位のstart reservationを取り、並行prepare/validate/apply/startを拒否する。Begin失敗時だけreservationを戻し、選択briefをretry可能に保つ。Begin成功後にStartedへ遷移する。
- start済みbriefは既存active draftの冪等retryには使えるが、`userRequestedNewWorkflow=true`の別資料には再利用できない。
- 監査snapshotへは`selection_source=agentDefault|userCard`だけを保存する。choice session、option、nonce、brief IDはjob payloadへ残さない。旧payloadでこのfieldが欠落した場合は`agentDefault`として読む。

### 完成sample thumbnail

- 外部bundleの`sample-thumbnails/<sample-id>.png`を任意登録できる。これはユーザーが方向を確認するためのUI artifactだけであり、modelやrendererの学習入力ではない。生成を拘束する正本はsample metadataとrecipeである。
- thumbnailはnon-interlaced PNGだけを受理する。JPEGはExif/XMP等の後置metadataを安全に検査し切れないため第2段階では拒否し、管理者がmetadataを除いたPNG derivativeへ変換する。
- path、symlink、magic、MIME、dimension、PNG chunk allowlist、全chunk CRC、完全なzlib scanline decodeをfail-closedに検証する。技術検査を権利・機密承認とは扱わない。
- 1画像256KiB・1600x1600以下・1.5Mpx以下、1profile 16画像・圧縮1MiB・6Mpx・展開32MiB以下、catalog全体は圧縮16MiB・48Mpx・展開128MiB以下とする。カードは最大3画像、raw合計240KiB、UI Resource全体384KiB以下とし、超える場合はthumbnailなしのmetadata表示へ落とす。
- `content_hash`はmanifest bytesに、sample ID/MIME/thumbnail SHA-256をsample ID順で合成する。thumbnail追加・変更時もprofile versionを上げ、atomic deploy後に再起動する。
- thumbnail catalogにはprofile単位ACLがない。対象deploymentの全PowerPoint利用者へ表示・配布してよい、非機密で権利確認済みのderivativeだけを置く。MCPのchoice TTLは、会話へ既に埋め込まれたdata URIをhostから削除しないため、保持・共有期間はhost policyにも従う。
- sampleと実生成の視覚一致はmetadata/recipeだけでは保証しない。代表slide previewとvisual regressionを後続段階で追加する。

## 却下した案

- 全依頼にカードを出す: 明確な依頼へ不要な往復を追加するため却下する。
- 候補ごとに50ページ分のbrief/Asset Planを複製する: schemaとtokenが案数に比例するため却下する。
- clientからstyle/brief/actionを受け取る: tamperとreplayの境界が広がるため却下する。
- model-facing fallbackへoption対応表を出す: UIを見ていないAgentが利用者選択を偽装できるため却下する。
- `pptx-mcp`がURLからsampleや画像を取得する: networkと権利境界を壊すため却下する。
- JPEGをmagic/dimension検査だけで許可する: scan後metadataを見落とすため却下する。
- sample metadataだけをcarouselへ出す: 完成形を見せる目的が弱いため、承認済みPNGを任意追加できる設計にする。ただし未登録時は正直にmetadata fallbackと表示する。

## トレードオフと残存課題

- clear pathのruntime tool callと人間往復は増えないが、prepare/apply/cancelのtool schemaとserver instructionsはtool-capable turnの恒常tokenを増やす。2026-08-11のUTF-8実測（description + input schema）はvalidate 7,698 bytes、prepare 9,638 bytes、apply 1,107 bytes、cancel 549 bytesで、Phase 2追加3 toolの合計は11,294 bytesである。各絶対値と追加合計を回帰testで上限固定する。
- 選択stateは単一processのmemory内にある。再起動またはnon-stickyなmulti-replicaではapplyがnot-foundとしてfail-closedになる。永続・共有storeを導入するまではsingle replicaまたはsticky routingを運用条件とする。
- 期限切れsessionのbounded tombstoneを保持しないため、prune後の古いactionはexpiredではなくnot-foundになる場合がある。生成は開始できないが、分類と長期replay監査は後続課題である。
- UI Resource自体は固定intentをpostするだけであり、hostのgenericなaction-to-prompt変換を認可境界にしない。対応host integrationはexactなintent／params allowlistを検証し、利用者へ保存・表示する本文と、server固定のmodel-only selection contextを分離する。`pptx-mcp`はhostで検証済みという主張も信用せず、apply時にopaqueな2 ID、利用者、会話、有効期限、profile hash、stateを必ず再検証する。stable MCP Appsのpredeclared resource、app visibility、JSON-RPC actionをhostが正式対応した時点で移行を再評価する。
- prepare後にresource構築が失敗した場合はpendingをrollbackする。表示後にhost側だけでmarkerが失われた場合は、利用者の明示または非対応判定後にcancelで復旧する。専用refetchはない。
- 引数なしcancelはcaller／conversation／pending stateをserverで検証するが、「hostでカードが表示できなかった」「利用者がsafe defaultを明示した」という実行理由までは検証・永続監査できない。この条件はAgent／host policyに依存し、job監査へ残るのはcancel後に直接検証したBriefの`selection_source=agentDefault`だけである。理由別監査が必要になった時点で、自由文を受けない列挙reasonと監査eventを別途設計する。
- default template IDはprofileへ束縛するが、template PPTX本体bytesをprofile content hashへ含めていない。template差替えはversion bumpとatomic deployを必須運用とし、bytes hash束縛は後続課題とする。
- `DesignBriefService`へcandidate validation、派生、choice state、reservationが集中した。Asset Brokerまたは代表previewを追加する前に、clock-awareな`DesignBriefSelectionStore`へ状態遷移を分離し、Pending→Selected→Starting→Started、pending-only replace/cancel、abort、expiry/restartを表駆動testへ移す。
- apply後の誤clickを取り消すcancel/reprepareは第2段階では実装しない。方針確定の監査とstart競合を崩さず補償操作を設計する必要がある。
- profile/sample単位ACL、asset approval registry、真のphoto/no-photo比較、代表slide先行previewは後続段階とする。

## 検証

- 候補最小2・最大3、no-photo時style別案最大1、null/重複/見た目が同じ候補の拒否をtestする。
- owner、期限、profile hash、option membership、same-option idempotency、different-option replay、stale actionをtestする。
- pendingがdirect/optional startを塞ぎ、apply後は選択briefだけを通し、start reservationの成功・失敗・並行実行をtestする。
- cancelのno-arg schema、no pending、cross-user/conversation非干渉、applyとのrace、selected/reserved/started拒否をtestする。
- HTML escaping、固定intent、CSP script hash、height-only resize、ID/URL/path/内部binding非露出、明色primaryのforeground contrastをtestする。
- PNG metadata/CRC/symlink/MIME/dimension/圧縮・pixel・展開量上限、thumbnailを含むcontent hash、card embedとmetadata fallbackをtestする。
- job auditの`userCard`伝播、opaque ID非保存、旧payloadの`agentDefault`互換をtestする。
