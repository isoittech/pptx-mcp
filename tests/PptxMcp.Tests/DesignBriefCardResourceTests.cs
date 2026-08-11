using System.Security.Cryptography;
using System.Net;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using ModelContextProtocol.Protocol;
using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class DesignBriefCardResourceTests
{
    private static readonly CallerContext Caller = new("user-a", "conversation-a", "message-a");

    [Fact]
    public void ResourceUsesFixedIntentOpaqueIdsStrictCspAndCompactProjection()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        DesignBriefSelectionTests.AddAlternativeDirection(root, materiallyDifferent: true);
        try
        {
            var (card, _) = PrepareCard(root);
            var result = DesignBriefCardResource.Create(card);
            var text = Assert.IsType<TextContentBlock>(result.Content[0]).Text;
            var embedded = Assert.IsType<EmbeddedResourceBlock>(result.Content[1]);
            var resource = Assert.IsType<TextResourceContents>(embedded.Resource);
            var html = resource.Text;

            Assert.False(result.IsError);
            Assert.Equal($"ui://pptx/design-brief/{card.ChoiceSessionId}", resource.Uri);
            Assert.Equal("text/html", resource.MimeType);
            Assert.True(Encoding.UTF8.GetByteCount(html) <= DesignBriefCardResource.MaximumResourceBytes);
            Assert.Contains("intent:'pptx.designBrief.select'", html, StringComparison.Ordinal);
            Assert.Contains("choiceSessionId:button.dataset.session", html, StringComparison.Ordinal);
            Assert.Contains("optionId:button.dataset.option", html, StringComparison.Ordinal);
            Assert.Contains($"data-session=\"{card.ChoiceSessionId}\"", html, StringComparison.Ordinal);
            Assert.Contains($"data-option=\"{card.Recommended.OptionId}\"", html, StringComparison.Ordinal);
            Assert.DoesNotContain(card.Recommended.Binding.BriefId, html, StringComparison.Ordinal);
            Assert.DoesNotContain("approved-general", html, StringComparison.Ordinal);
            Assert.DoesNotContain("prohibited_rules", html, StringComparison.Ordinal);
            Assert.DoesNotContain("template_source", html, StringComparison.Ordinal);
            Assert.DoesNotContain("http://", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("https://", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("default-src 'none'", html, StringComparison.Ordinal);
            Assert.Contains("connect-src 'none'", html, StringComparison.Ordinal);
            Assert.DoesNotContain("script-src 'unsafe-inline'", html, StringComparison.Ordinal);
            Assert.Contains("script-src 'sha256-", html, StringComparison.Ordinal);
            Assert.Contains("Math.min(900,Math.max(240", html, StringComparison.Ordinal);
            Assert.DoesNotContain("scrollWidth", html, StringComparison.Ordinal);
            Assert.Contains("主な前提", html, StringComparison.Ordinal);
            Assert.Contains("素材方針: 承認済み素材のみ", html, StringComparison.Ordinal);
            Assert.Contains(HtmlEncoder.Default.Encode("推奨案: Standard"), html, StringComparison.Ordinal);
            Assert.Contains(HtmlEncoder.Default.Encode("別案: Report"), html, StringComparison.Ordinal);
            Assert.DoesNotContain(card.ChoiceSessionId, text, StringComparison.Ordinal);
            Assert.DoesNotContain(card.Recommended.OptionId, text, StringComparison.Ordinal);
            Assert.DoesNotContain(card.Recommended.Binding.BriefId, text, StringComparison.Ordinal);
            Assert.Contains("end this turn", text, StringComparison.OrdinalIgnoreCase);
            var defensiveSerialization = JsonSerializer.Serialize(card);
            Assert.DoesNotContain(card.Recommended.Binding.BriefId, defensiveSerialization, StringComparison.Ordinal);
            Assert.DoesNotContain("BrandProfile", defensiveSerialization, StringComparison.Ordinal);

            var script = Regex.Match(html, "<script>(?<body>.*?)</script>", RegexOptions.Singleline);
            var hash = Regex.Match(html, "script-src 'sha256-(?<hash>[^']+)'", RegexOptions.Singleline);
            Assert.True(script.Success);
            Assert.True(hash.Success);
            Assert.Equal(
                Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(script.Groups["body"].Value))),
                WebUtility.HtmlDecode(hash.Groups["hash"].Value));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ResourceEscapesUserAndProfileTextAndDerivesReadableActionForegrounds()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        DesignBriefSelectionTests.AddAlternativeDirection(root, materiallyDifferent: true);
        SetBrightCardPalette(root);
        try
        {
            var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true);
            var catalog = new BrandProfileCatalog(options);
            var service = new DesignBriefService(options, TimeProvider.System, catalog);
            var (baseBrief, assetPlan) = BrandProfileTestFactory.CreateBrief(
                DesignBriefSelectionTests.GetReference(catalog));
            var brief = baseBrief with
            {
                Audience = "<img src=x onerror=alert('audience')>",
                Assumptions =
                [
                    new DesignAssumption(
                        "<b onmouseover=alert('assumption')>",
                        DesignAssumptionStatus.Inferred),
                ],
            };
            var card = service.Prepare(
                Caller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null);
            var result = DesignBriefCardResource.Create(card);
            var html = Assert.IsType<TextResourceContents>(
                Assert.IsType<EmbeddedResourceBlock>(result.Content[1]).Resource).Text;

            Assert.DoesNotContain("<img src=x", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<b onmouseover", html, StringComparison.Ordinal);
            Assert.Contains("&lt;img", html, StringComparison.Ordinal);
            Assert.Contains("&lt;b", html, StringComparison.Ordinal);
            Assert.Contains("--on-primary:#000000", html, StringComparison.Ordinal);
            Assert.Contains("--on-accent:#000000", html, StringComparison.Ordinal);
            Assert.Contains(".pill strong{color:var(--text)}", html, StringComparison.Ordinal);
            Assert.Contains(".meta{margin-top:9px;font-size:11px;color:var(--muted)}", html, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void UnregisteredOptionSampleUsesHonestMetadataFallback()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        DesignBriefSelectionTests.AddAlternativeDirection(root, materiallyDifferent: true);
        RemoveAlternativeSamples(root);
        try
        {
            var (card, _) = PrepareCard(root);
            var html = Assert.IsType<TextResourceContents>(
                Assert.IsType<EmbeddedResourceBlock>(
                    DesignBriefCardResource.Create(card).Content[1]).Resource).Text;

            Assert.Contains(HtmlEncoder.Default.Encode("別案: Report"), html, StringComparison.Ordinal);
            Assert.Contains("実行レシピに紐づく完成サンプルは未登録", html, StringComparison.Ordinal);
            Assert.Contains("metadata fallback", html, StringComparison.Ordinal);
            Assert.DoesNotContain("KPI alias", html, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ResourceEmbedsOnlyCatalogValidatedPngThumbnailsWithinTheCardCap()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        DesignBriefSelectionTests.AddAlternativeDirection(root, materiallyDifferent: true);
        var thumbnails = Directory.CreateDirectory(Path.Combine(root, "general", "sample-thumbnails"));
        try
        {
            File.WriteAllBytes(
                Path.Combine(thumbnails.FullName, "sample-cover-low.png"),
                BrandSampleThumbnailLoaderTests.CreatePng(2, 1, 8, 2, 0x11));
            File.WriteAllBytes(
                Path.Combine(thumbnails.FullName, "sample-cover-report.png"),
                BrandSampleThumbnailLoaderTests.CreatePng(2, 1, 8, 2, 0x22));

            var (card, _) = PrepareCard(root);
            var result = DesignBriefCardResource.Create(card);
            var text = Assert.IsType<TextContentBlock>(result.Content[0]).Text;
            var html = Assert.IsType<TextResourceContents>(
                Assert.IsType<EmbeddedResourceBlock>(result.Content[1]).Resource).Text;

            Assert.Equal(2, Regex.Matches(html, "data:image/png;base64,").Count);
            Assert.True(Encoding.UTF8.GetByteCount(html) <= DesignBriefCardResource.MaximumResourceBytes);
            Assert.DoesNotContain("sample-cover-low", text, StringComparison.Ordinal);
            Assert.DoesNotContain(thumbnails.FullName, html, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private static (PreparedDesignBriefCard Card, DesignBriefService Service) PrepareCard(string root)
    {
        var options = BrandProfileTestFactory.CreateOptions(root, requireDesignBrief: true);
        var catalog = new BrandProfileCatalog(options);
        var service = new DesignBriefService(options, TimeProvider.System, catalog);
        var (brief, assetPlan) = BrandProfileTestFactory.CreateBrief(
            DesignBriefSelectionTests.GetReference(catalog));
        return (
            service.Prepare(
                Caller,
                brief,
                assetPlan,
                [new DesignBriefStyleAlternative(
                    "report",
                    "detailed",
                    ["cover-report-airy", "kpi-report-detailed"])],
                null),
            service);
    }

    private static void SetBrightCardPalette(string root)
    {
        var path = Path.Combine(root, "general", "brand-profile.json");
        var manifest = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        var roles = manifest["color_roles"]!.AsObject();
        roles["primary"] = "#FFFF00";
        roles["accent"] = "#F5E942";
        roles["background"] = "#111111";
        roles["surface"] = "#111111";
        roles["text"] = "#FFFFFF";
        roles["muted_text"] = "#B8C0C8";
        File.WriteAllText(path, manifest.ToJsonString());
    }

    private static void RemoveAlternativeSamples(string root)
    {
        var path = Path.Combine(root, "general", "brand-profile.json");
        var manifest = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        foreach (var recipe in manifest["layout_recipes"]!.AsArray()
                     .Select(static item => item!.AsObject())
                     .Where(static item => item["style_direction_id"]!.GetValue<string>() == "report"))
        {
            recipe["sample_ids"] = new JsonArray();
        }

        var filtered = manifest["samples"]!.AsArray()
            .Where(static item => item!["style_direction_id"]!.GetValue<string>() != "report")
            .Select(static item => item!.DeepClone())
            .ToArray();
        manifest["samples"] = new JsonArray(filtered);
        File.WriteAllText(path, manifest.ToJsonString());
    }
}
