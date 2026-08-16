using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;
using PptxMcp.Domain;

namespace PptxMcp.Storage;

public sealed class VisualObjectAssetRepository(TimeProvider timeProvider)
{
    public const int MaximumBatchObjects = 8;
    public const int MaximumObjectsPerSlide = 3;
    public const int MaximumConversationObjects = 24;
    private const int MaximumActiveObjects = 512;
    private static readonly TimeSpan AssetLifetime = TimeSpan.FromHours(1);
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web);
    private readonly ConcurrentDictionary<string, VisualObjectAssetManifest> assets = new(StringComparer.Ordinal);
    private readonly object prepareLock = new();

    public VisualObjectPreparationView Prepare(CallerContext caller, IReadOnlyList<VisualObjectBrief> briefs)
    {
        if (briefs is null || briefs.Count is < 1 or > MaximumBatchObjects || briefs.Any(static item => item is null))
        {
            throw new PptxValidationException(
                "visual_object_batch_invalid",
                $"Provide between 1 and {MaximumBatchObjects} complete visual object briefs in one call.");
        }

        ValidateBriefs(briefs);
        lock (prepareLock)
        {
            PruneExpired();
            var activeForConversation = assets.Values.Count(item =>
                string.Equals(item.UserScope, caller.UserScope, StringComparison.Ordinal)
                && string.Equals(item.ConversationScope, caller.ConversationScope, StringComparison.Ordinal));
            if (activeForConversation + briefs.Count > MaximumConversationObjects || assets.Count + briefs.Count > MaximumActiveObjects)
            {
                throw new PptxValidationException(
                    "visual_object_capacity_reached",
                    $"The visual object limit is {MaximumConversationObjects} per conversation. Reuse prepared objects or wait for expiry; do not retry in a loop.");
            }

            var now = timeProvider.GetUtcNow();
            var views = new List<VisualObjectAssetView>(briefs.Count);
            foreach (var brief in briefs)
            {
                var id = Guid.NewGuid().ToString("N");
                var manifest = new VisualObjectAssetManifest(
                    id,
                    caller.UserScope,
                    caller.ConversationScope,
                    brief,
                    Fingerprint(brief),
                    now,
                    now.Add(AssetLifetime));
                if (!assets.TryAdd(id, manifest))
                {
                    throw new InvalidOperationException("Could not allocate a visual object asset ID.");
                }

                views.Add(CreateView(manifest));
            }

            return new VisualObjectPreparationView(
                "ready",
                views,
                "Copy each opaque asset_id only to the planned slide's visualObjects list in this conversation. Use at most three per slide. The renderer will create editable native PowerPoint shapes; do not convert IDs into SVG, XML, coordinates, URLs, or paths.");
        }
    }

    public VisualObjectAssetManifest GetOwned(CallerContext caller, string assetId) =>
        GetOwnedByScope(caller.UserScope, caller.ConversationScope, assetId);

    internal VisualObjectAssetManifest GetOwnedByScope(
        string userScope,
        string conversationScope,
        string assetId)
    {
        if (!IsAssetId(assetId)
            || !assets.TryGetValue(assetId, out var manifest)
            || !string.Equals(manifest.UserScope, userScope, StringComparison.Ordinal)
            || !string.Equals(manifest.ConversationScope, conversationScope, StringComparison.Ordinal)
            || manifest.ExpiresAt <= timeProvider.GetUtcNow())
        {
            throw NotFound();
        }

        return manifest;
    }

    private static void ValidateBriefs(IReadOnlyList<VisualObjectBrief> briefs)
    {
        foreach (var group in briefs.GroupBy(static item => item.SlideNumber))
        {
            if (group.Key is < 1 or > 50 || group.Count() > MaximumObjectsPerSlide)
            {
                throw new PptxValidationException(
                    "visual_object_slide_limit_invalid",
                    $"slideNumber must be 1-50 and each slide may prepare at most {MaximumObjectsPerSlide} objects.");
            }

            if (group.Count(static item => item.Emphasis == VisualObjectEmphasis.Strong) > 1)
            {
                throw new PptxValidationException(
                    "visual_object_emphasis_invalid",
                    "Use at most one strong visual object per slide; keep the remaining objects subtle or standard.");
            }
        }

        for (var index = 0; index < briefs.Count; index++)
        {
            var item = briefs[index];
            if (item.Label is not null
                && (string.IsNullOrWhiteSpace(item.Label)
                    || item.Label.Length > 48
                    || item.Label.Any(static character => char.IsControl(character))))
            {
                throw new PptxValidationException(
                    "visual_object_label_invalid",
                    $"objects[{index}].label must be 1-48 characters without control characters when provided.");
            }

            var compatible = item.Archetype switch
            {
                VisualObjectArchetype.Arrow => item.VisualPurpose is VisualObjectPurpose.Direction or VisualObjectPurpose.Growth,
                VisualObjectArchetype.CurvedArrow => item.VisualPurpose is VisualObjectPurpose.Direction or VisualObjectPurpose.Cycle,
                VisualObjectArchetype.Frame => item.VisualPurpose is VisualObjectPurpose.Emphasis or VisualObjectPurpose.Grouping,
                VisualObjectArchetype.Callout => item.VisualPurpose is VisualObjectPurpose.Annotation or VisualObjectPurpose.Emphasis,
                VisualObjectArchetype.Bracket => item.VisualPurpose == VisualObjectPurpose.Grouping,
                VisualObjectArchetype.Ring => item.VisualPurpose is VisualObjectPurpose.Cycle or VisualObjectPurpose.Emphasis,
                VisualObjectArchetype.Ribbon => item.VisualPurpose is VisualObjectPurpose.Emphasis or VisualObjectPurpose.Annotation,
                _ => false,
            };
            if (!compatible)
            {
                throw new PptxValidationException(
                    "visual_object_semantics_invalid",
                    $"objects[{index}] uses an archetype that does not express its visualPurpose. Choose the semantic native shape that matches the message.");
            }

            if (item.Archetype == VisualObjectArchetype.CurvedArrow
                && item.Orientation != VisualObjectOrientation.Clockwise)
            {
                throw new PptxValidationException(
                    "visual_object_orientation_invalid",
                    $"objects[{index}].orientation must be clockwise for curvedArrow.");
            }
        }
    }

    private void PruneExpired()
    {
        var now = timeProvider.GetUtcNow();
        foreach (var item in assets)
        {
            if (item.Value.ExpiresAt <= now)
            {
                assets.TryRemove(item.Key, out _);
            }
        }
    }

    private static string Fingerprint(VisualObjectBrief brief)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(brief, SerializerOptions);
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    private static bool IsAssetId(string value) =>
        value.Length == 32 && value.All(static character => character is >= '0' and <= '9' or >= 'a' and <= 'f');

    private static VisualObjectAssetView CreateView(VisualObjectAssetManifest manifest) =>
        new(
            manifest.AssetId,
            manifest.Brief.SlideNumber,
            manifest.Brief.Archetype,
            manifest.Brief.PlacementRole,
            manifest.Brief.Style,
            manifest.Brief.Emphasis,
            $"{manifest.Brief.Style} {manifest.Brief.Archetype} · {manifest.Brief.Emphasis} · {manifest.Brief.PlacementRole}",
            manifest.ExpiresAt);

    private static PptxValidationException NotFound() =>
        new("visual_object_asset_not_found", "The visual object asset was not found for this user and conversation, or it expired.");
}
