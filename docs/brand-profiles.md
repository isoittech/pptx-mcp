# Brand Profile bundle

Brand Profileは導入環境が管理する外部設定です。会社名、実テンプレート、ロゴ、社内URL、素材、認証情報をOSSリポジトリへ追加せず、`BrandProfilesRoot`へ読み取り専用でマウントしてください。

## ディレクトリ

```text
<BrandProfilesRoot>/
└── general/
    └── brand-profile.json
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

`semantic_kind`、`design_style`、`density`、`motif`、`theme_preset`、`variant`は現行Visual Deckレンダラーが実装する値だけを受け付けます。variantは意味layoutと一致するだけでなく、slide追加時に内容件数等の実装条件も検証されます。`Metrics`の`spotlight`は縦方向の重なりを防ぐため正確に3件だけを受理し、recipeにも`item_count: 3`を必須とします。第1段階の`item_count`はこの組合せ以外には指定できません。

## sampleと素材の第1段階

`samples`は完成サンプルを検索・選択するための要約metadataです。第1段階のMCPはsample PPTXやthumbnailを返さず、画像の取得・挿入も行いません。`required_asset_roles`を持つrecipeはcatalogへ登録できますが、外部画像を挿入できる安全なfile ID経路が実装されるまでは、Design Briefにそのrecipeを確定できません。`nativeDraw`または画像なしで成立する別recipeを選んでください。

Asset Planで素材を使わないページを表す正規形は、`preferred_medium=none`、`acquisition=none`、`fallback=none`、`status=omitted`、`license_status=notRequired`です。asset metadataは省略し、`required_asset_roles`が空のrecipeを選びます。`noAssetLayout`は「素材なし」の別名ではありません。`userUpload`または`approvedLibrary`を計画したものの、第1段階では挿入せず画像なしrecipeへ切り替える項目に限り、`status=fallbackSelected`と組み合わせて使います。

## 反映と更新

起動時に全manifestを検証し、プロセス内では検証済みsnapshotを保持します。manifestの生バイト列から計算したSHA-256が`content_hash`です。更新時は宣言`version`も変更してサーバーを再起動し、Agentにcatalogを再取得させてDesign Briefを再検証してください。

無引数のcatalog呼出しは常にsummaryだけです。`purpose`、`density`、`styleDirectionId`で絞る場合も、先に一覧から選んだ正確な`profileId`が必要です。これにより複数profileのrule、recipe、sampleを1応答へ展開しません。

生成後はprofile version/hash、style direction、各slideのrecipe契約と、検証済みDesign Briefの最小監査snapshotをjob payloadへ残します。監査snapshotにはsource policy、確定・推定assumption、Asset Planの意味tokenとopaque参照だけを保存し、期限付き`brief_id`、URL、path、認証情報は保存しません。refineでもこのbindingを保持し、追加ページの新しいAsset Planを検証できない第1段階ではprofile-bound deckへのinsertを拒否します。
