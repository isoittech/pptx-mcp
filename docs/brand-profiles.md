# Brand Profile bundle

Brand Profileは導入環境が管理する外部設定です。会社名、実テンプレート、ロゴ、社内URL、素材、認証情報をOSSリポジトリへ追加せず、`BrandProfilesRoot`へ読み取り専用でマウントしてください。

## ディレクトリ

```text
<BrandProfilesRoot>/
└── general/
    ├── brand-profile.json
    └── sample-thumbnails/
        └── <sample-id>.png
```

bundleディレクトリ名とmanifestの`id`は完全一致させます。ID、recipe、sample、style direction、素材collectionの参照は、英数字、ハイフン、アンダースコアからなる最大128文字のopaque IDです。URLやパスではありません。

## schema version 1

次は企業固有値を含まない最小例です。

```json
{
  "schema_version": 1,
  "id": "general",
  "version": "1.0.0",
  "name": "General presentation profile",
  "description": "A generic profile for concise business presentations.",
  "template_source": "none",
  "template_id": "",
  "color_roles": {
    "primary": "#17324D",
    "secondary": "#49657D",
    "accent": "#D96C36",
    "background": "#F7F8FA",
    "surface": "#FFFFFF",
    "text": "#17212B",
    "muted_text": "#66717C",
    "positive": "#2B7A4B",
    "warning": "#A35B00",
    "critical": "#B42318",
    "data_series": ["#17324D", "#D96C36", "#49657D"]
  },
  "typography": {
    "heading_font": "Aptos Display",
    "body_font": "Aptos"
  },
  "voice_rules": [
    "Use concise conclusion-led headings."
  ],
  "visual_rules": {
    "photography": [],
    "illustration": [],
    "iconography": ["Use a restrained native icon set."],
    "native_shapes": ["Prefer editable diagrams for explanatory content."],
    "tables": ["Keep headers distinct and rows scannable."],
    "charts": ["Highlight only the decision-relevant series."],
    "backgrounds": ["Reserve dark backgrounds for section messages."],
    "emphasis": ["Use accent color sparingly."]
  },
  "visual_object_policy": {
    "allowed_archetypes": ["arrow", "curvedArrow", "frame", "callout", "bracket", "ring", "ribbon"],
    "allowed_styles": ["quietCorporate", "editorial", "technical"],
    "default_style": "quietCorporate",
    "maximum_per_slide": 3,
    "maximum_per_deck": 16,
    "strong_requires_focal_purpose": true
  },
  "prohibited_rules": [
    "Do not use decorative placeholders as finished visuals."
  ],
  "requires_confirmation_rules": [
    "Confirm the use of identifiable people."
  ],
  "approved_asset_collection_ids": ["approved-general"],
  "style_directions": [
    {
      "id": "standard",
      "name": "Standard",
      "summary": "Conclusion-led business slides with restrained visual emphasis.",
      "recommended_for": ["cover", "kpi"],
      "design_style": "executive",
      "default_density": "balanced",
      "supported_densities": ["airy", "balanced", "detailed"],
      "motif": "geometric",
      "theme_preset": "minimal"
    }
  ],
  "layout_recipes": [
    {
      "id": "cover-airy",
      "purpose": "cover",
      "semantic_kind": "Title",
      "variant": "auto",
      "density": "airy",
      "style_direction_id": "standard",
      "required_asset_roles": [],
      "sample_ids": ["sample-cover-low"]
    },
    {
      "id": "kpi-balanced",
      "purpose": "kpi",
      "semantic_kind": "Metrics",
      "variant": "spotlight",
      "density": "balanced",
      "style_direction_id": "standard",
      "required_asset_roles": [],
      "sample_ids": ["sample-kpi-medium"],
      "item_count": 3
    }
  ],
  "samples": [
    {
      "id": "sample-cover-low",
      "title": "Short cover",
      "summary": "Low-information cover emphasizing one message.",
      "purpose": "cover",
      "density": "airy",
      "style_direction_id": "standard",
      "recipe_id": "cover-airy",
      "information_level": "low"
    },
    {
      "id": "sample-kpi-medium",
      "title": "KPI summary",
      "summary": "Medium-information KPI composition with three key values.",
      "purpose": "kpi",
      "density": "balanced",
      "style_direction_id": "standard",
      "recipe_id": "kpi-balanced",
      "information_level": "medium"
    }
  ]
}
```

`template_source`は`default`または`none`です。`default`では`template_id`を導入環境の`DefaultTemplateId`と完全一致させ、`none`では空文字にします。profileから任意テンプレートパスやアップロードIDを指定することはできません。

manifestは256KB以下、profileは最大32件です。style directionは最大8件、recipeは最大64件、sampleは最大96件です。`data_series`は1〜8色を指定します。textとbackground/surfaceは4.5:1以上、muted textとbackground/surfaceは3.0:1以上のコントラストが必要です。

`semantic_kind`、`design_style`、`density`、`motif`、`theme_preset`、`variant`は現行Visual Deckレンダラーが実装する値だけを受け付けます。variantは意味layoutと一致するだけでなく、slide追加時に内容件数等の実装条件も検証されます。`Metrics`の`spotlight`は縦方向の重なりを防ぐため正確に3件だけを受理し、recipeにも`item_count: 3`を必須とします。Processの`loop`、Timeline/Roadmapの`stepped`、Funnelの`pyramid`は3〜6件だけを受理します。`NativeDiagram` recipeはvariant=`auto`とし、実slideのdiagram kind/node/edge上限を別途検証します。第1段階の`item_count`はMetrics spotlight以外には指定できません。

`visual_object_policy`は導入環境が使える補助図形と強さを制限します。server capは1ページ3件、1デッキ24件で、profileはそれ以下にだけ狭められます。`strong_requires_focal_purpose=true`ではstrong objectを`visualPurpose=emphasis`だけに限定します。ブランド準拠は図形を増やすことではありません。通常は`quietCorporate`と`subtle`を既定とし、矢印は方向・成長・循環、枠／吹き出しは1つの判断や数値を示す場合だけ使ってください。

style directionの`theme_preset`はrenderer対応値として必須ですが、現schemaはprofileの`color_roles`を全方向へ明示してpresetの基礎paletteを上書きします。したがってpreset名だけを変えても実効palette差にはなりません。方向差は`design_style`、`motif`、用途別recipeの構造差で作り、方向別paletteは将来のschema拡張までprofile内で表現できると誤認しないでください。

## sample thumbnailと素材

`samples`は完成サンプルを検索・選択するための要約metadataです。任意で`sample-thumbnails/<sample-id>.png`を置くと、条件付きDesign Briefカードのcarouselへ表示できます。thumbnailは利用者が方向を確認するためのUI artifactだけで、modelやrendererの学習入力ではありません。生成を拘束する正本はsample metadataとrecipeです。代表sampleと実生成の一致は、後続の代表slide previewとvisual regressionで確認してください。

thumbnailはmetadataを持たないnon-interlaced PNG derivativeだけを使います。JPEGは受理しません。path、symlink、PNG signature、dimension、chunk allowlist、全chunk CRC、完全なzlib scanline decodeを起動時に検査します。1画像は256KiB・1600x1600・1.5Mpx以下、1profileは16画像・圧縮1MiB・6Mpx・展開32MiB以下、catalog全体は圧縮16MiB・48Mpx・展開128MiB以下です。複数sampleを登録する通常運用では、16:9画像を960x540程度に抑えることを推奨します。たとえば1280x720を9枚置くと8.29Mpxとなり、profile上限を超えます。カードは最大3画像、raw合計240KiB、UI Resource全体384KiB以下とし、収まらない場合は画像なしのmetadata表示へ落とします。

この技術検査は権利・機密承認ではありません。現catalogにprofile単位ACLはないため、対象deploymentの全PowerPoint利用者へ表示・配布してよい、非機密で権利確認済みのderivativeだけを登録してください。data URIは会話hostへ渡った後、Design Briefの有効期限だけでは削除されません。hostの保持・共有policyも満たす必要があります。元PPTX、人物・製品画像、社内限定sampleを無条件にthumbnail化しないでください。

sample thumbnailはPPTXへ挿入せず、model/rendererの学習入力にも使いません。Webや承認済みライブラリからの画像取得も行いません。一方、[ADR 0016](adr/0016-conversation-scoped-image-assets-and-media-split.md)以降は、LibreChatへユーザーが添付したJPEG/PNGだけを`pptx_register_uploaded_image_asset`で会話scope付き・metadataなしPNGへ登録できます。登録済み`asset_id`を使う`Media/split` recipeは`required_asset_roles`を1件以上持たせ、Asset Planを`userUpload`、`ready`、`userProvided`、`fallback=none`、crop、`text_safe_area=left|right`で確定します。

Asset Planで素材を使わないページを表す正規形は、`preferred_medium=none`、`acquisition=none`、`fallback=none`、`status=omitted`、`license_status=notRequired`です。asset metadataは省略し、`required_asset_roles`が空のrecipeを選びます。`noAssetLayout`は「素材なし」の別名ではありません。未登録`userUpload`または現時点で解決未対応の`approvedLibrary`を画像なしrecipeへ切り替える項目に限り、`status=fallbackSelected`と組み合わせます。登録済み`userUpload`を実際に使う場合は`fallback=none`です。

最小の画像recipe例は次の通りです。会社名、内部URL、実asset IDはbundleへ保存しません。

```json
{
  "id": "product-media-balanced",
  "purpose": "product",
  "semantic_kind": "Media",
  "variant": "split",
  "density": "balanced",
  "style_direction_id": "standard",
  "required_asset_roles": ["hero_image"],
  "sample_ids": []
}
```

## 反映と更新

起動時に全manifestとthumbnailを検証し、プロセス内では読み込んだbyte snapshotを保持します。`content_hash`はmanifest bytesと、sample ID順のsample ID・MIME・thumbnail SHA-256を合成して計算します。thumbnailの追加・差替えでも宣言`version`を変更し、bundleをatomicに配備してサーバーを再起動し、Agentにcatalogを再取得させてDesign Briefを再検証してください。

`template_source=default`はtemplate IDを固定しますが、現段階のprofile hashは既定template PPTX本体のbytesを含みません。template差替え時もversionを上げ、profile bundleとtemplateをatomicに配備してください。

無引数のcatalog呼出しはsummaryとcompactな`style_directions`候補だけです。そこで正確な`profileId`と`styleDirectionId`を選び、2回目かつ最後の呼出しでdetail、recipe、sample要約を取得します。単一用途だけ`purpose`／`density`も指定でき、複数用途は方向だけで全用途recipeを取得します。これにより複数profileのrule、recipe、sampleや、全方向のrecipeを1応答へ展開しません。

生成後はprofile version/hash、style direction、各slideのrecipe契約と、検証済みDesign Briefの最小監査snapshotをjob payloadへ残します。監査snapshotにはsource policy、確定・推定assumption、Asset Planの意味token、`selection_source=agentDefault|userCard`、opaque参照だけを保存し、期限付き`brief_id`、choice session、option、URL、path、認証情報は保存しません。refineでもこのbindingを保持し、prepared visual object参照は置換slideで省略された場合に元ページからmaterializeします。明示された異なるIDは拒否します。追加ページの新しいAsset Planを検証できない第2段階ではprofile-bound deckへのinsertを拒否します。
