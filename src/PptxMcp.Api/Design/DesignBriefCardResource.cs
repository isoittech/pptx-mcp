using System.Text;
using System.Text.Encodings.Web;
using System.Security.Cryptography;
using ModelContextProtocol.Protocol;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Design;

internal static class DesignBriefCardResource
{
    internal const int MaximumResourceBytes = 384 * 1024;
    private const int MaximumSamples = 3;
    private const int MaximumEmbeddedThumbnailBytes = 240 * 1024;
    private const string ClientScript = """(()=>{const report=()=>parent.postMessage({type:'ui-size-change',payload:{height:Math.min(900,Math.max(240,document.documentElement.scrollHeight))}},'*');const send=(button)=>{if(button.disabled)return;document.querySelectorAll('[data-option]').forEach(item=>item.disabled=true);parent.postMessage({type:'intent',payload:{intent:'pptx.designBrief.select',params:{choiceSessionId:button.dataset.session,optionId:button.dataset.option}}},'*');};document.querySelectorAll('[data-option]').forEach(button=>button.addEventListener('click',()=>send(button)));const toggle=document.querySelector('.toggle');const alternatives=document.querySelector('.alternatives');if(toggle&&alternatives)toggle.addEventListener('click',()=>{const open=alternatives.classList.toggle('open');toggle.setAttribute('aria-expanded',String(open));report();});const slides=[...document.querySelectorAll('.sample')];let index=0;const counter=document.querySelector('.sample-counter');const show=(next)=>{if(!slides.length)return;index=(next+slides.length)%slides.length;slides.forEach((slide,i)=>slide.classList.toggle('active',i===index));if(counter)counter.textContent=`${index+1} / ${slides.length}`;report();};document.querySelector('.sample-prev')?.addEventListener('click',()=>show(index-1));document.querySelector('.sample-next')?.addEventListener('click',()=>show(index+1));new ResizeObserver(report).observe(document.body);show(0);report();})();""";

    public static CallToolResult Create(PreparedDesignBriefCard card)
    {
        ArgumentNullException.ThrowIfNull(card);
        var samples = SelectSamples(card);
        var html = BuildHtml(card, samples, includeThumbnails: true);
        if (Encoding.UTF8.GetByteCount(html) > MaximumResourceBytes)
        {
            html = BuildHtml(card, samples, includeThumbnails: false);
        }

        if (Encoding.UTF8.GetByteCount(html) > MaximumResourceBytes)
        {
            throw new PptxValidationException(
                "design_brief_ui_resource_too_large",
                $"The Design Brief UI Resource exceeds the {MaximumResourceBytes}-byte limit even without thumbnails.");
        }

        return new CallToolResult
        {
            Content =
            [
                new TextContentBlock
                {
                    Text = "A Design Brief choice card is ready. This card has not authorized PowerPoint generation. Present the embedded UI Resource marker, end this turn, and wait for the user. After the fixed pptx.designBrief.select intent arrives, call pptx_apply_design_brief_action with only its choiceSessionId and optionId. Do not call pptx_validate_design_brief or pptx_start_visual_deck while this card is pending. If the host cannot render the card, fail closed; use pptx_cancel_design_brief_selection only to recover from the invisible unselected card, then validate the safe recommendation directly.",
                },
                new EmbeddedResourceBlock
                {
                    Resource = new TextResourceContents
                    {
                        Uri = $"ui://pptx/design-brief/{card.ChoiceSessionId}",
                        MimeType = "text/html",
                        Text = html,
                    },
                },
            ],
            IsError = false,
        };
    }

    public static CallToolResult CreateError(string tool, PptxValidationException exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        var instruction = exception.Code switch
        {
            "design_brief_choice_pending" =>
                "Render the existing card and wait for its user action. Replace it only after the user explicitly asks.",
            "brand_profile_version_mismatch" =>
                "Refresh pptx_get_design_catalog and prepare a new card from the current immutable profile version. Do not reuse the stale profile hash.",
            "design_brief_card_choice_required"
                or "design_brief_alternative_has_no_visual_difference"
                or "design_brief_no_photo_has_no_visual_difference" =>
                "Skip the card and use pptx_validate_design_brief, or provide a genuinely distinct executable choice.",
            "design_brief_start_in_progress" =>
                "Wait for the current Visual Deck start call; do not prepare another card concurrently.",
            "design_brief_choice_already_applied" =>
                "Start only the briefId already returned by pptx_apply_design_brief_action.",
            "design_brief_capacity_reached" or "design_brief_user_capacity_reached" =>
                "Stop retrying. Wait for an existing brief or card to complete or expire before preparing another card.",
            "design_brief_confirmation_required" or "design_brief_questions_unresolved" =>
                "Ask only the material unresolved questions or record a safe explicit fallback before preparing the finalized card.",
            "design_brief_ui_resource_too_large" =>
                "Do not retry the same card. Use pptx_validate_design_brief for the safe recommendation and report that the comparison card could not be rendered.",
            _ =>
                "Correct the invalid field without adding URLs, paths, or unverified assets, then retry once.",
        };
        return new CallToolResult
        {
            Content =
            [
                new TextContentBlock
                {
                    Text = $"invalid_input\ntool={tool}\ncode={exception.Code}\nmessage={exception.Message}\ninstruction={instruction}",
                },
            ],
            IsError = true,
        };
    }

    private static string BuildHtml(
        PreparedDesignBriefCard card,
        IReadOnlyList<CardSample> samples,
        bool includeThumbnails)
    {
        var recommended = card.Recommended;
        var brief = recommended.Binding.Brief;
        var profile = recommended.Binding.Profile;
        var direction = recommended.Binding.StyleDirection;
        var colors = profile.Detail.ColorRoles;
        var onPrimary = ReadableForeground(colors.Primary);
        var onAccent = ReadableForeground(colors.Accent);
        var scriptHash = Convert.ToBase64String(
            SHA256.HashData(Encoding.UTF8.GetBytes(ClientScript)));
        var builder = new StringBuilder(24 * 1024);
        builder.Append("<!doctype html><html lang=\"ja\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; base-uri 'none'; form-action 'none'; connect-src 'none'; img-src data:; media-src 'none'; object-src 'none'; font-src 'none'; style-src 'unsafe-inline'; script-src 'sha256-");
        builder.Append(E(scriptHash));
        builder.Append("""
            '">
              <style>
                *{box-sizing:border-box}html,body{margin:0;padding:0;background:transparent;color:var(--text);font-family:system-ui,-apple-system,"Segoe UI",sans-serif}body{padding:8px}.brief{--primary:
            """);
        builder.Append(E(colors.Primary));
        builder.Append(";--accent:");
        builder.Append(E(colors.Accent));
        builder.Append(";--background:");
        builder.Append(E(colors.Background));
        builder.Append(";--surface:");
        builder.Append(E(colors.Surface));
        builder.Append(";--text:");
        builder.Append(E(colors.Text));
        builder.Append(";--muted:");
        builder.Append(E(colors.MutedText));
        builder.Append(";--on-primary:");
        builder.Append(onPrimary);
        builder.Append(";--on-accent:");
        builder.Append(onAccent);
        builder.Append("""
            ;max-width:920px;border:1px solid color-mix(in srgb,var(--primary) 18%,#d6dbe1);border-radius:18px;overflow:hidden;background:var(--surface);box-shadow:0 12px 32px rgba(16,24,40,.10)}.head{padding:22px 24px 18px;background:var(--primary);color:var(--on-primary)}.kicker{font-size:12px;font-weight:750;letter-spacing:.08em;text-transform:uppercase}.head h1{font-size:23px;line-height:1.3;margin:7px 0 5px}.head p{margin:0;font-size:14px;line-height:1.55}.body{padding:20px 24px 24px;background:color-mix(in srgb,var(--background) 55%,var(--surface))}.facts{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:9px;margin-bottom:16px}.fact{min-height:76px;padding:11px 12px;border:1px solid color-mix(in srgb,var(--primary) 13%,#d6dbe1);border-radius:11px;background:var(--surface)}.fact b{display:block;margin-bottom:4px;color:var(--muted);font-size:11px}.fact span{font-size:14px;font-weight:700;line-height:1.35}.strategy{padding:14px 15px;border-left:4px solid var(--accent);border-radius:8px;background:color-mix(in srgb,var(--accent) 8%,var(--surface));font-size:14px;line-height:1.55}.assumptions{margin:10px 0;padding:9px 12px;border:1px solid color-mix(in srgb,var(--primary) 14%,#d6dbe1);border-radius:9px;background:var(--surface);font-size:12px}.assumptions summary{cursor:pointer;font-weight:700}.assumptions ul{margin:8px 0 2px;padding-left:20px}.assumptions li{margin:5px 0;line-height:1.4}.assumptions em{color:var(--muted);font-style:normal}.policy{margin-left:8px;color:var(--muted);font-weight:400}.assets{display:flex;flex-wrap:wrap;gap:7px;margin:15px 0}.pill{padding:6px 9px;border-radius:999px;background:var(--surface);border:1px solid color-mix(in srgb,var(--primary) 14%,#d6dbe1);font-size:12px}.pill strong{color:var(--text)}.section-title{margin:20px 0 10px;font-size:14px}.samples{position:relative}.sample{display:none;grid-template-columns:minmax(190px,42%) 1fr;gap:14px;align-items:stretch;border:1px solid color-mix(in srgb,var(--primary) 14%,#d6dbe1);border-radius:13px;background:var(--surface);overflow:hidden}.sample.active{display:grid}.thumb{min-height:142px;background:var(--primary);display:flex;align-items:center;justify-content:center;overflow:hidden}.thumb img{width:100%;height:100%;min-height:142px;object-fit:cover}.mock{width:84%;aspect-ratio:16/9;padding:10px;background:var(--background);border-radius:5px;box-shadow:0 7px 18px rgba(0,0,0,.2)}.mock:before{content:"";display:block;width:46%;height:7px;background:var(--primary);border-radius:4px;margin-bottom:10px}.mock:after{content:"サムネイル未登録";display:grid;place-items:center;height:62%;border:1px dashed color-mix(in srgb,var(--primary) 35%,#fff);color:var(--muted);font-size:10px}.sample-copy{padding:14px 16px}.sample-copy b{display:block;font-size:15px;margin-bottom:5px}.sample-copy p{margin:0;color:var(--muted);font-size:13px;line-height:1.5}.meta{margin-top:9px;font-size:11px;color:var(--muted)}.sample-nav{display:flex;align-items:center;justify-content:flex-end;gap:7px;margin-top:8px}.sample-nav button,.toggle{border:1px solid color-mix(in srgb,var(--primary) 25%,#d6dbe1);border-radius:8px;background:var(--surface);color:var(--text);padding:6px 10px;cursor:pointer}.choices{display:flex;flex-wrap:wrap;gap:9px;margin-top:18px}.choose{border:0;border-radius:10px;padding:11px 15px;font-weight:750;cursor:pointer;background:var(--primary);color:var(--on-primary)}.choose.secondary{background:var(--surface);color:var(--text);border:1px solid color-mix(in srgb,var(--primary) 30%,#d6dbe1)}.choose.photo-free{background:var(--accent);color:var(--on-accent)}button:focus-visible{outline:3px solid var(--accent);outline-offset:2px}button[disabled]{cursor:not-allowed;opacity:.55}.alternatives{display:none;margin-top:12px;padding:12px;border-radius:10px;background:var(--surface);border:1px solid color-mix(in srgb,var(--primary) 14%,#d6dbe1)}.alternatives.open{display:block}.alt{display:flex;justify-content:space-between;gap:12px;align-items:center;padding:9px 0}.alt+.alt{border-top:1px solid #e7e9ed}.alt-copy b{display:block;font-size:13px}.alt-copy span{display:block;margin-top:3px;color:var(--muted);font-size:12px}.notice{margin:14px 0 0;color:var(--muted);font-size:11px;line-height:1.45}@media(max-width:680px){.facts{grid-template-columns:repeat(2,minmax(0,1fr))}.sample.active{grid-template-columns:1fr}.thumb{min-height:118px}.body{padding:16px}.head{padding:18px}.choices{display:grid}.choose{width:100%}}
              </style>
            </head>
            <body>
              <main class="brief" aria-labelledby="brief-title">
                <header class="head"><div class="kicker">Design Brief</div><h1 id="brief-title">
            """);
        builder.Append(E(direction.Name));
        builder.Append(" を推奨します</h1><p>");
        builder.Append(E(direction.Summary));
        builder.Append("</p></header><section class=\"body\"><div class=\"facts\">");
        AppendFact(builder, "目的", brief.Purpose);
        AppendFact(builder, "読者", brief.Audience);
        AppendFact(builder, "利用場面", DeliveryLabel(brief.DeliveryMode));
        AppendFact(builder, "情報密度", DensityLabel(brief.Density));
        builder.Append("</div><div class=\"strategy\"><strong>見せ方:</strong> ");
        builder.Append(E(brief.VisualStrategy));
        builder.Append("<br><strong>トーン:</strong> ");
        builder.Append(E(brief.DesiredTone));
        builder.Append("</div>");
        AppendAssumptions(builder, brief);
        builder.Append("<div class=\"assets\" aria-label=\"素材計画の状態\">");
        AppendPill(builder, "準備済み", recommended.AssetSummary.ReadyCount);
        AppendPill(builder, "安全な代替へ確定", recommended.AssetSummary.FallbackSelectedCount);
        AppendPill(builder, "素材なし", recommended.AssetSummary.OmittedCount);
        builder.Append("</div>");
        if (card.NoPhoto is not null)
        {
            builder.Append("<div class=\"strategy\"><strong>画像案について:</strong> 現段階では推奨案も外部画像を挿入せず、確定済みの安全な代替で生成します。画像を使わない別構成は、レシピ自体を図解中心または素材なしへ切り替えます。</div>");
        }

        AppendSamples(builder, samples, includeThumbnails);
        builder.Append("<div class=\"choices\">");
        AppendChoiceButton(
            builder,
            card.ChoiceSessionId,
            recommended.OptionId,
            "推奨案で進む",
            "choose");
        if (card.Alternatives.Count > 0)
        {
            builder.Append("<button class=\"toggle\" type=\"button\" aria-expanded=\"false\" aria-controls=\"alternative-list\">別案を見る</button>");
        }

        if (card.NoPhoto is not null)
        {
            AppendChoiceButton(
                builder,
                card.ChoiceSessionId,
                card.NoPhoto.OptionId,
                "画像を使わない別構成",
                "choose photo-free");
        }

        builder.Append("</div>");
        if (card.Alternatives.Count > 0)
        {
            builder.Append("<div id=\"alternative-list\" class=\"alternatives\">");
            foreach (var alternative in card.Alternatives)
            {
                builder.Append("<div class=\"alt\"><div class=\"alt-copy\"><b>");
                builder.Append(E(alternative.Binding.StyleDirection.Name));
                builder.Append("</b><span>");
                builder.Append(E(alternative.Binding.StyleDirection.Summary));
                builder.Append(" · ");
                builder.Append(E(DensityLabel(alternative.Binding.Brief.Density)));
                builder.Append("</span></div>");
                AppendChoiceButton(
                    builder,
                    card.ChoiceSessionId,
                    alternative.OptionId,
                    "この別案を選ぶ",
                    "choose secondary");
                builder.Append("</div>");
            }

            builder.Append("</div>");
        }

        builder.Append("<p class=\"notice\">選択するとこの方針を確定し、生成を開始します。別案への変更はできないため、内容を確認して選んでください。ボタンにはサーバー発行の不透明な選択IDだけが含まれます。有効期限: ");
        builder.Append(E(card.ExpiresAt.ToString("u")));
        builder.Append("</p></section></main>");
        builder.Append("<script>");
        builder.Append(ClientScript);
        builder.Append("</script></body></html>");
        return builder.ToString();
    }

    private static List<CardSample> SelectSamples(PreparedDesignBriefCard card)
    {
        var result = new List<CardSample>(MaximumSamples);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var options = new List<PreparedDesignBriefOption> { card.Recommended };
        options.AddRange(card.Alternatives);
        if (card.NoPhoto is not null)
        {
            options.Add(card.NoPhoto);
        }

        foreach (var option in options)
        {
            var recipeIds = option.Binding.AssetPlan.Values
                .Select(static item => item.RecipeId)
                .ToHashSet(StringComparer.Ordinal);
            var sample = option.Binding.Profile.Samples.FirstOrDefault(item =>
                recipeIds.Contains(item.RecipeId)
                && seen.Add(item.Id));
            BrandSampleThumbnail? thumbnail = null;
            if (sample is not null)
            {
                option.Binding.Profile.SampleThumbnails.TryGetValue(sample.Id, out thumbnail);
            }

            result.Add(new CardSample(OptionLabel(option), sample, thumbnail));
            if (result.Count == MaximumSamples)
            {
                break;
            }
        }

        return result;
    }

    private static void AppendSamples(
        StringBuilder builder,
        IReadOnlyList<CardSample> samples,
        bool includeThumbnails)
    {
        builder.Append("<h2 class=\"section-title\">完成サンプル</h2>");
        if (samples.Count == 0)
        {
            builder.Append("<div class=\"strategy\">この方向の完成サンプルは未登録です。レシピとブランドルールだけで判断します。</div>");
            return;
        }

        var embeddedBytes = 0;
        builder.Append("<div class=\"samples\" aria-live=\"polite\">");
        for (var index = 0; index < samples.Count; index++)
        {
            var sample = samples[index];
            builder.Append(index == 0 ? "<article class=\"sample active\">" : "<article class=\"sample\">");
            builder.Append("<div class=\"thumb\">");
            var canEmbed = includeThumbnails
                && sample.Summary is not null
                && sample.Thumbnail is not null
                && embeddedBytes + sample.Thumbnail.Bytes.Length <= MaximumEmbeddedThumbnailBytes;
            if (canEmbed)
            {
                embeddedBytes += sample.Thumbnail!.Bytes.Length;
                builder.Append("<img alt=\"");
                builder.Append(E($"完成サンプル: {sample.Summary!.Title}"));
                builder.Append("\" src=\"data:");
                builder.Append(E(sample.Thumbnail.MimeType));
                builder.Append(";base64,");
                builder.Append(Convert.ToBase64String(sample.Thumbnail.Bytes.Span));
                builder.Append("\">");
            }
            else
            {
                builder.Append("<div class=\"mock\" role=\"img\" aria-label=\"サムネイル未登録。メタデータによるサンプル案内\"></div>");
            }

            builder.Append("</div><div class=\"sample-copy\"><b>");
            builder.Append(E(sample.OptionLabel));
            if (sample.Summary is null)
            {
                builder.Append("</b><p>この選択肢の実行レシピに紐づく完成サンプルは未登録です。ブランドルールとレシピ情報で判断します。</p><div class=\"meta\">metadata fallback</div>");
            }
            else
            {
                builder.Append(" — ");
                builder.Append(E(sample.Summary.Title));
                builder.Append("</b><p>");
                builder.Append(E(sample.Summary.Summary));
                builder.Append("</p><div class=\"meta\">");
                builder.Append(E($"{sample.Summary.Purpose} · {DensityLabel(sample.Summary.Density)} · 情報量 {InformationLabel(sample.Summary.InformationLevel)}"));
                builder.Append("</div>");
            }
            builder.Append("</div></article>");
        }

        builder.Append("</div>");
        if (samples.Count > 1)
        {
            builder.Append("<div class=\"sample-nav\"><button class=\"sample-prev\" type=\"button\" aria-label=\"前のサンプル\">←</button><span class=\"sample-counter\">1 / ");
            builder.Append(samples.Count);
            builder.Append("</span><button class=\"sample-next\" type=\"button\" aria-label=\"次のサンプル\">→</button></div>");
        }
    }

    private static void AppendFact(StringBuilder builder, string label, string value)
    {
        builder.Append("<div class=\"fact\"><b>");
        builder.Append(E(label));
        builder.Append("</b><span>");
        builder.Append(E(value));
        builder.Append("</span></div>");
    }

    private static void AppendAssumptions(StringBuilder builder, DesignBriefSpec brief)
    {
        var shown = brief.Assumptions
            .OrderBy(static item => item.Status == DesignAssumptionStatus.Inferred ? 0 : 1)
            .Take(3)
            .ToArray();
        builder.Append("<details class=\"assumptions\"><summary>主な前提 <span class=\"policy\">素材方針: ");
        builder.Append(brief.SourcePolicy == DesignSourcePolicy.ApprovedOnly
            ? "承認済み素材のみ"
            : "承認済み・ユーザー提供素材");
        builder.Append("</span></summary><ul>");
        foreach (var assumption in shown)
        {
            builder.Append("<li><em>");
            builder.Append(assumption.Status == DesignAssumptionStatus.Inferred ? "推定" : "確認済み");
            builder.Append(":</em> ");
            builder.Append(E(assumption.Text));
            builder.Append("</li>");
        }

        if (brief.Assumptions.Count > shown.Length)
        {
            builder.Append("<li><em>ほか ");
            builder.Append(brief.Assumptions.Count - shown.Length);
            builder.Append(" 件</em></li>");
        }

        builder.Append("</ul></details>");
    }

    private static void AppendPill(StringBuilder builder, string label, int count)
    {
        builder.Append("<span class=\"pill\">");
        builder.Append(E(label));
        builder.Append(" <strong>");
        builder.Append(count);
        builder.Append("</strong></span>");
    }

    private static void AppendChoiceButton(
        StringBuilder builder,
        string choiceSessionId,
        string optionId,
        string label,
        string cssClass)
    {
        builder.Append("<button type=\"button\" class=\"");
        builder.Append(E(cssClass));
        builder.Append("\" data-session=\"");
        builder.Append(E(choiceSessionId));
        builder.Append("\" data-option=\"");
        builder.Append(E(optionId));
        builder.Append("\">");
        builder.Append(E(label));
        builder.Append("</button>");
    }

    private static string DeliveryLabel(DesignDeliveryMode deliveryMode) => deliveryMode switch
    {
        DesignDeliveryMode.Projection => "投影",
        DesignDeliveryMode.Handout => "配布",
        DesignDeliveryMode.Both => "投影・配布",
        _ => "未指定",
    };

    private static string DensityLabel(string density) => density.ToLowerInvariant() switch
    {
        "airy" => "少なめ",
        "balanced" => "標準",
        "detailed" => "多め",
        _ => density,
    };

    private static string InformationLabel(string informationLevel) => informationLevel.ToLowerInvariant() switch
    {
        "low" => "少",
        "medium" => "中",
        "high" => "多",
        _ => informationLevel,
    };

    private static string E(string value) => HtmlEncoder.Default.Encode(value);

    private static string ReadableForeground(string background)
    {
        var red = Convert.ToInt32(background.Substring(1, 2), 16) / 255d;
        var green = Convert.ToInt32(background.Substring(3, 2), 16) / 255d;
        var blue = Convert.ToInt32(background.Substring(5, 2), 16) / 255d;
        static double Linearize(double value) => value <= 0.04045
            ? value / 12.92
            : Math.Pow((value + 0.055) / 1.055, 2.4);
        var luminance = 0.2126 * Linearize(red)
            + 0.7152 * Linearize(green)
            + 0.0722 * Linearize(blue);
        var blackContrast = (luminance + 0.05) / 0.05;
        var whiteContrast = 1.05 / (luminance + 0.05);
        return blackContrast >= whiteContrast ? "#000000" : "#FFFFFF";
    }

    private static string OptionLabel(PreparedDesignBriefOption option) => option.Kind switch
    {
        DesignBriefCardOptionKind.Recommended => $"推奨案: {option.Binding.StyleDirection.Name}",
        DesignBriefCardOptionKind.Alternative => $"別案: {option.Binding.StyleDirection.Name}",
        DesignBriefCardOptionKind.NoPhoto => "画像を使わない別構成",
        _ => "選択案",
    };

    private sealed record CardSample(
        string OptionLabel,
        BrandSampleSummary? Summary,
        BrandSampleThumbnail? Thumbnail);
}
