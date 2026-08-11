using System.Collections.ObjectModel;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Design;

public sealed partial class BrandProfileCatalog
{
    private const int MaximumManifestBytes = 256 * 1024;
    private const int MaximumProfiles = 32;
    private const int MaximumStyleDirections = 8;
    private const int MaximumRecipes = 64;
    private const int MaximumSamples = 96;
    private static readonly HashSet<string> SupportedTemplateSources = new(StringComparer.OrdinalIgnoreCase)
    {
        "default",
        "none",
    };

    private static readonly HashSet<string> SupportedDesignStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        "executive",
        "editorial",
        "bold",
        "technical",
        "playful",
    };

    private static readonly HashSet<string> SupportedDensities = new(StringComparer.OrdinalIgnoreCase)
    {
        "airy",
        "balanced",
        "detailed",
    };

    private static readonly HashSet<string> SupportedMotifs = new(StringComparer.OrdinalIgnoreCase)
    {
        "geometric",
        "orbit",
        "nodes",
        "ribbon",
        "none",
    };

    private static readonly HashSet<string> SupportedThemes = new(StringComparer.OrdinalIgnoreCase)
    {
        "midnight",
        "aurora",
        "sunset",
        "forest",
        "minimal",
        "ocean",
        "berry",
        "clay",
        "cyber",
    };

    private static readonly HashSet<string> SupportedInformationLevels = new(StringComparer.OrdinalIgnoreCase)
    {
        "low",
        "medium",
        "high",
    };

    private static readonly JsonSerializerOptions ManifestSerializerOptions = new(JsonSerializerDefaults.Web)
    {
        MaxDepth = 32,
        PropertyNameCaseInsensitive = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
    };

    private readonly PptxMcpOptions options;
    private readonly Lazy<IReadOnlyDictionary<string, BrandProfileSnapshot>> profiles;

    public BrandProfileCatalog(IOptions<PptxMcpOptions> options)
    {
        this.options = options.Value;
        profiles = new Lazy<IReadOnlyDictionary<string, BrandProfileSnapshot>>(
            LoadProfiles,
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public void EnsureReady() => _ = profiles.Value;

    public DesignCatalogView Query(
        string? profileId = null,
        string? purpose = null,
        string? density = null,
        string? styleDirectionId = null)
    {
        ValidateOptionalIdentifier(profileId, "profileId");
        ValidateOptionalIdentifier(purpose, "purpose");
        ValidateOptionalIdentifier(styleDirectionId, "styleDirectionId");
        if (profileId is null
            && (purpose is not null || density is not null || styleDirectionId is not null))
        {
            throw new PptxValidationException(
                "design_catalog_profile_required",
                "purpose, density, and styleDirectionId filters require one exact profileId from the summary-only catalog response.");
        }

        if (density is not null && !SupportedDensities.Contains(density))
        {
            throw new PptxValidationException(
                "design_catalog_density_invalid",
                "density must be airy, balanced, or detailed.");
        }

        var snapshot = profiles.Value;
        if (snapshot.Count == 0)
        {
            return new DesignCatalogView(
                "not_configured",
                [],
                "No external Brand Profile is configured. Continue with the ordinary Visual Deck workflow unless this deployment requires a Design Brief.");
        }

        IEnumerable<BrandProfileSnapshot> selected = snapshot.Values;
        if (profileId is not null)
        {
            if (!snapshot.TryGetValue(profileId, out var exactProfile))
            {
                throw new PptxValidationException(
                    "brand_profile_not_found",
                    "The requested Brand Profile ID was not found. Call pptx_get_design_catalog without filters and select an exact returned ID.");
            }

            selected = [exactProfile];
        }

        var includeDetails = profileId is not null;
        var views = selected
            .OrderBy(static profile => profile.Id, StringComparer.Ordinal)
            .Select(profile => CreateView(profile, includeDetails, purpose, density, styleDirectionId))
            .Where(view => !includeDetails || view.Recipes.Count > 0 || profileId is not null)
            .ToArray();

        if (styleDirectionId is not null
            && profileId is not null
            && views[0].Detail!.StyleDirections.All(direction =>
                !string.Equals(direction.Id, styleDirectionId, StringComparison.Ordinal)))
        {
            throw new PptxValidationException(
                "brand_style_direction_not_found",
                "styleDirectionId was not found in the selected Brand Profile.");
        }

        return new DesignCatalogView(
            "available",
            views,
            includeDetails
                ? "Select exact recipe IDs for every slide, then call pptx_validate_design_brief with the immutable profile version and content_hash shown here."
                : "Call pptx_get_design_catalog again with profileId and optional purpose, density, or styleDirectionId filters to retrieve compact recipes and sample guidance.");
    }

    public BrandProfileSnapshot GetSnapshot(BrandProfileReference reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        ValidateIdentifier(reference.Id, "brand_profile.id");
        ValidateVersion(reference.Version);
        if (!ContentHashRegex().IsMatch(reference.ContentHash))
        {
            throw new PptxValidationException(
                "brand_profile_hash_invalid",
                "brand_profile.content_hash must be the exact lowercase SHA-256 hash returned by pptx_get_design_catalog.");
        }

        if (!profiles.Value.TryGetValue(reference.Id, out var profile))
        {
            throw new PptxValidationException(
                "brand_profile_not_found",
                "The selected Brand Profile is not available in this deployment.");
        }

        if (!string.Equals(profile.Version, reference.Version, StringComparison.Ordinal)
            || !string.Equals(profile.ContentHash, reference.ContentHash, StringComparison.Ordinal))
        {
            throw new PptxValidationException(
                "brand_profile_version_mismatch",
                "The Brand Profile version or content_hash changed. Refresh pptx_get_design_catalog and validate the Design Brief again.");
        }

        return profile;
    }

    private ReadOnlyDictionary<string, BrandProfileSnapshot> LoadProfiles()
    {
        var root = Path.GetFullPath(options.BrandProfilesRoot);
        var rootDirectory = new DirectoryInfo(root);
        if (!rootDirectory.Exists)
        {
            if (options.RequireDesignBrief)
            {
                throw new InvalidOperationException(
                    "PptxMcp:RequireDesignBrief is enabled but BrandProfilesRoot does not exist.");
            }

            return new ReadOnlyDictionary<string, BrandProfileSnapshot>(
                new Dictionary<string, BrandProfileSnapshot>(StringComparer.Ordinal));
        }

        if (rootDirectory.LinkTarget is not null)
        {
            throw new InvalidOperationException("PptxMcp:BrandProfilesRoot must not be a symbolic link.");
        }

        var loaded = new Dictionary<string, BrandProfileSnapshot>(StringComparer.Ordinal);
        foreach (var parent in rootDirectory
                     .EnumerateDirectories("*", SearchOption.TopDirectoryOnly)
                     .OrderBy(static directory => directory.Name, StringComparer.Ordinal))
        {
            if (!SafeIdentifierRegex().IsMatch(parent.Name) || parent.LinkTarget is not null)
            {
                throw new InvalidOperationException(
                    "Brand Profile bundle directories must use safe, non-symbolic identifiers.");
            }

            var manifestFile = new FileInfo(Path.Combine(parent.FullName, "brand-profile.json"));
            if (!manifestFile.Exists)
            {
                continue;
            }

            if (loaded.Count >= MaximumProfiles)
            {
                throw new InvalidOperationException($"Brand Profile catalog must not exceed {MaximumProfiles} profiles.");
            }

            if (manifestFile.LinkTarget is not null)
            {
                throw new InvalidOperationException("Brand Profile bundle names and files must be safe, non-symbolic identifiers.");
            }

            if (manifestFile.Length is <= 0 or > MaximumManifestBytes)
            {
                throw new InvalidOperationException(
                    $"Brand Profile manifest must contain 1 to {MaximumManifestBytes} bytes.");
            }

            var bytes = File.ReadAllBytes(manifestFile.FullName);
            BrandProfileManifest manifest;
            try
            {
                manifest = JsonSerializer.Deserialize<BrandProfileManifest>(bytes, ManifestSerializerOptions)
                    ?? throw new JsonException("Manifest was empty.");
            }
            catch (JsonException exception)
            {
                throw new InvalidOperationException("Brand Profile manifest is invalid JSON or contains unsupported fields.", exception);
            }

            var contentHash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
            var profile = ValidateAndCreateSnapshot(manifest, parent.Name, contentHash);
            if (!loaded.TryAdd(profile.Id, profile))
            {
                throw new InvalidOperationException("Brand Profile IDs must be unique.");
            }
        }

        if (options.RequireDesignBrief && loaded.Count == 0)
        {
            throw new InvalidOperationException(
                "PptxMcp:RequireDesignBrief is enabled but no valid Brand Profile bundle is configured.");
        }

        return new ReadOnlyDictionary<string, BrandProfileSnapshot>(loaded);
    }

    private BrandProfileSnapshot ValidateAndCreateSnapshot(
        BrandProfileManifest manifest,
        string bundleId,
        string contentHash)
    {
        if (manifest.SchemaVersion != 1)
        {
            throw new InvalidOperationException("Brand Profile schema_version must be 1.");
        }

        ValidateIdentifier(manifest.Id, "id");
        if (!string.Equals(manifest.Id, bundleId, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Brand Profile id must exactly match its bundle directory name.");
        }

        ValidateVersion(manifest.Version);
        ValidateText(manifest.Name, "name", 1, 100);
        ValidateText(manifest.Description, "description", 1, 400);
        if (!SupportedTemplateSources.Contains(manifest.TemplateSource))
        {
            throw new InvalidOperationException("Brand Profile template_source must be default or none.");
        }

        if (string.Equals(manifest.TemplateSource, "default", StringComparison.OrdinalIgnoreCase))
        {
            ValidateIdentifier(manifest.TemplateId, "template_id");
            if (string.IsNullOrWhiteSpace(options.DefaultTemplateId)
                || !string.Equals(manifest.TemplateId, options.DefaultTemplateId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "A Brand Profile using template_source=default must name the deployment DefaultTemplateId exactly.");
            }
        }
        else if (!string.IsNullOrEmpty(manifest.TemplateId))
        {
            throw new InvalidOperationException("template_id must be empty when template_source=none.");
        }

        ValidateColorRoles(manifest.ColorRoles);
        ValidateTypography(manifest.Typography);
        ValidateTextList(manifest.VoiceRules, "voice_rules", 1, 16, 300);
        ValidateVisualRules(manifest.VisualRules);
        ValidateTextList(manifest.ProhibitedRules, "prohibited_rules", 0, 24, 300);
        ValidateTextList(manifest.RequiresConfirmationRules, "requires_confirmation_rules", 0, 24, 300);
        ValidateIdentifierList(
            manifest.ApprovedAssetCollectionIds,
            "approved_asset_collection_ids",
            maximumCount: 32);

        if (manifest.StyleDirections.Count is < 1 or > MaximumStyleDirections)
        {
            throw new InvalidOperationException(
                $"style_directions must contain between 1 and {MaximumStyleDirections} items.");
        }

        var styleDirections = manifest.StyleDirections.ToArray();
        EnsureUnique(styleDirections.Select(static item => item.Id), "style direction");
        foreach (var direction in styleDirections)
        {
            ValidateStyleDirection(direction);
        }

        if (manifest.LayoutRecipes.Count is < 1 or > MaximumRecipes)
        {
            throw new InvalidOperationException(
                $"layout_recipes must contain between 1 and {MaximumRecipes} items.");
        }

        var recipes = manifest.LayoutRecipes.ToArray();
        EnsureUnique(recipes.Select(static item => item.Id), "layout recipe");
        foreach (var recipe in recipes)
        {
            ValidateRecipe(recipe, styleDirections);
        }

        if (manifest.Samples.Count > MaximumSamples)
        {
            throw new InvalidOperationException($"samples must not contain more than {MaximumSamples} items.");
        }

        var samples = manifest.Samples.ToArray();
        EnsureUnique(samples.Select(static item => item.Id), "sample");
        foreach (var sample in samples)
        {
            ValidateSample(sample, recipes);
        }

        foreach (var recipe in recipes)
        {
            ValidateIdentifierList(recipe.RequiredAssetRoles, $"layout_recipes[{recipe.Id}].required_asset_roles", 12);
            ValidateIdentifierList(recipe.SampleIds, $"layout_recipes[{recipe.Id}].sample_ids", 12);
            if (recipe.SampleIds.Any(sampleId => samples.All(sample => !string.Equals(sample.Id, sampleId, StringComparison.Ordinal))))
            {
                throw new InvalidOperationException($"Layout recipe {recipe.Id} references an unknown sample ID.");
            }
        }

        var summary = new BrandProfileCatalogSummary(
            manifest.Id,
            manifest.Name,
            manifest.Version,
            contentHash,
            manifest.Description,
            manifest.TemplateSource.ToLowerInvariant(),
            styleDirections.Select(static direction => direction.Id).ToArray());
        var detail = new BrandProfileCatalogDetail(
            summary,
            manifest.ColorRoles,
            manifest.Typography,
            manifest.VoiceRules.ToArray(),
            manifest.VisualRules,
            manifest.ProhibitedRules.ToArray(),
            manifest.RequiresConfirmationRules.ToArray(),
            manifest.ApprovedAssetCollectionIds.ToArray(),
            styleDirections);
        return new BrandProfileSnapshot(detail, manifest.TemplateId, recipes, samples);
    }

    private static DesignCatalogProfileView CreateView(
        BrandProfileSnapshot profile,
        bool includeDetails,
        string? purpose,
        string? density,
        string? styleDirectionId)
    {
        if (!includeDetails)
        {
            return new DesignCatalogProfileView(profile.Detail.Summary, null, [], []);
        }

        var recipes = profile.LayoutRecipes
            .Where(recipe => purpose is null || string.Equals(recipe.Purpose, purpose, StringComparison.OrdinalIgnoreCase))
            .Where(recipe => density is null || string.Equals(recipe.Density, density, StringComparison.OrdinalIgnoreCase))
            .Where(recipe => styleDirectionId is null || string.Equals(recipe.StyleDirectionId, styleDirectionId, StringComparison.Ordinal))
            .ToArray();
        var recipeIds = recipes.Select(static recipe => recipe.Id).ToHashSet(StringComparer.Ordinal);
        var samples = profile.Samples
            .Where(sample => recipeIds.Contains(sample.RecipeId))
            .ToArray();
        var details = profile.Detail with
        {
            StyleDirections = styleDirectionId is null
                ? profile.Detail.StyleDirections
                : profile.Detail.StyleDirections
                    .Where(direction => string.Equals(direction.Id, styleDirectionId, StringComparison.Ordinal))
                    .ToArray(),
        };
        return new DesignCatalogProfileView(profile.Detail.Summary, details, recipes, samples);
    }

    private static void ValidateStyleDirection(BrandStyleDirection direction)
    {
        ValidateIdentifier(direction.Id, "style_directions.id");
        ValidateText(direction.Name, "style_directions.name", 1, 100);
        ValidateText(direction.Summary, "style_directions.summary", 1, 300);
        ValidateIdentifierList(direction.RecommendedFor, "style_directions.recommended_for", 16);
        if (!SupportedDesignStyles.Contains(direction.DesignStyle))
        {
            throw new InvalidOperationException("style_directions.design_style is unsupported by the renderer.");
        }

        if (!SupportedDensities.Contains(direction.DefaultDensity))
        {
            throw new InvalidOperationException("style_directions.default_density is unsupported by the renderer.");
        }

        if (direction.SupportedDensities.Count is < 1 or > 3
            || direction.SupportedDensities.Any(density => !SupportedDensities.Contains(density))
            || !direction.SupportedDensities.Contains(direction.DefaultDensity, StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "style_directions.supported_densities must contain one to three renderer densities and include default_density.");
        }

        if (!SupportedMotifs.Contains(direction.Motif))
        {
            throw new InvalidOperationException("style_directions.motif is unsupported by the renderer.");
        }

        if (!SupportedThemes.Contains(direction.ThemePreset))
        {
            throw new InvalidOperationException("style_directions.theme_preset is unsupported by the renderer.");
        }
    }

    private static void ValidateRecipe(
        BrandLayoutRecipe recipe,
        IReadOnlyList<BrandStyleDirection> styleDirections)
    {
        ValidateIdentifier(recipe.Id, "layout_recipes.id");
        ValidateIdentifier(recipe.Purpose, "layout_recipes.purpose");
        if (!SupportedDensities.Contains(recipe.Density))
        {
            throw new InvalidOperationException("layout_recipes.density is unsupported by the renderer.");
        }

        var styleDirection = styleDirections.SingleOrDefault(
            direction => string.Equals(direction.Id, recipe.StyleDirectionId, StringComparison.Ordinal));
        if (styleDirection is null)
        {
            throw new InvalidOperationException("layout_recipes.style_direction_id references an unknown style direction.");
        }

        if (!styleDirection.SupportedDensities.Contains(recipe.Density, StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "layout_recipes.density must be listed in the referenced style direction's supported_densities.");
        }

        var variantSupported = recipe.Variant.Equals("auto", StringComparison.OrdinalIgnoreCase)
            || recipe.Variant.Equals("split", StringComparison.OrdinalIgnoreCase) && recipe.SemanticKind == VisualSlideKind.Bullets
            || recipe.Variant.Equals("spotlight", StringComparison.OrdinalIgnoreCase)
                && recipe.SemanticKind is VisualSlideKind.Metrics or VisualSlideKind.Cards
            || recipe.Variant.Equals("editorial", StringComparison.OrdinalIgnoreCase)
                && recipe.SemanticKind == VisualSlideKind.StructuredBrief;
        if (!variantSupported)
        {
            throw new InvalidOperationException(
                "layout_recipes.variant must be implemented for the selected semantic_kind.");
        }

        if (recipe.ItemCount is <= 0 or > 10)
        {
            throw new InvalidOperationException("layout_recipes.item_count must be between 1 and 10 when specified.");
        }

        if (recipe.SemanticKind == VisualSlideKind.Metrics
            && recipe.Variant.Equals("spotlight", StringComparison.OrdinalIgnoreCase)
            && recipe.ItemCount != 3)
        {
            throw new InvalidOperationException(
                "A Metrics spotlight layout recipe must set item_count to exactly 3.");
        }

        if (recipe.ItemCount is not null
            && (recipe.SemanticKind != VisualSlideKind.Metrics
                || !recipe.Variant.Equals("spotlight", StringComparison.OrdinalIgnoreCase)))
        {
            throw new InvalidOperationException(
                "layout_recipes.item_count is currently supported only for Metrics spotlight recipes.");
        }
    }

    private static void ValidateSample(
        BrandSampleSummary sample,
        IReadOnlyList<BrandLayoutRecipe> recipes)
    {
        ValidateIdentifier(sample.Id, "samples.id");
        ValidateText(sample.Title, "samples.title", 1, 160);
        ValidateText(sample.Summary, "samples.summary", 1, 400);
        ValidateIdentifier(sample.Purpose, "samples.purpose");
        ValidateIdentifier(sample.StyleDirectionId, "samples.style_direction_id");
        ValidateIdentifier(sample.RecipeId, "samples.recipe_id");
        if (!SupportedDensities.Contains(sample.Density)
            || !SupportedInformationLevels.Contains(sample.InformationLevel))
        {
            throw new InvalidOperationException("samples density or information_level is invalid.");
        }

        var recipe = recipes.SingleOrDefault(recipe => string.Equals(recipe.Id, sample.RecipeId, StringComparison.Ordinal));
        if (recipe is null
            || !string.Equals(recipe.Purpose, sample.Purpose, StringComparison.Ordinal)
            || !string.Equals(recipe.Density, sample.Density, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(recipe.StyleDirectionId, sample.StyleDirectionId, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Each sample must match its referenced recipe purpose, density, and style direction.");
        }
    }

    private static void ValidateColorRoles(BrandColorRoles roles)
    {
        ArgumentNullException.ThrowIfNull(roles);
        foreach (var color in new[]
                 {
                     roles.Primary,
                     roles.Secondary,
                     roles.Accent,
                     roles.Background,
                     roles.Surface,
                     roles.Text,
                     roles.MutedText,
                     roles.Positive,
                     roles.Warning,
                     roles.Critical,
                 })
        {
            if (!ColorRegex().IsMatch(color))
            {
                throw new InvalidOperationException("Brand Profile color roles must use #RRGGBB values.");
            }
        }

        if (roles.DataSeries is null
            || roles.DataSeries.Count is < 1 or > 8
            || roles.DataSeries.Any(color => !ColorRegex().IsMatch(color)))
        {
            throw new InvalidOperationException(
                "Brand Profile color_roles.data_series must contain one to eight #RRGGBB values.");
        }

        EnsureContrast(roles.Text, roles.Background, 4.5, "text/background");
        EnsureContrast(roles.Text, roles.Surface, 4.5, "text/surface");
        EnsureContrast(roles.MutedText, roles.Background, 3.0, "muted_text/background");
        EnsureContrast(roles.MutedText, roles.Surface, 3.0, "muted_text/surface");
    }

    private static void ValidateTypography(BrandTypography typography)
    {
        ArgumentNullException.ThrowIfNull(typography);
        ValidateText(typography.HeadingFont, "typography.heading_font", 1, 80);
        ValidateText(typography.BodyFont, "typography.body_font", 1, 80);
    }

    private static void ValidateVisualRules(BrandVisualRuleSet rules)
    {
        ArgumentNullException.ThrowIfNull(rules);
        ValidateTextList(rules.Photography, "visual_rules.photography", 0, 16, 300);
        ValidateTextList(rules.Illustration, "visual_rules.illustration", 0, 16, 300);
        ValidateTextList(rules.Iconography, "visual_rules.iconography", 0, 16, 300);
        ValidateTextList(rules.NativeShapes, "visual_rules.native_shapes", 0, 16, 300);
        ValidateTextList(rules.Tables, "visual_rules.tables", 0, 16, 300);
        ValidateTextList(rules.Charts, "visual_rules.charts", 0, 16, 300);
        ValidateTextList(rules.Backgrounds, "visual_rules.backgrounds", 0, 16, 300);
        ValidateTextList(rules.Emphasis, "visual_rules.emphasis", 0, 16, 300);
    }

    private static void ValidateTextList(
        IReadOnlyList<string> values,
        string field,
        int minimumCount,
        int maximumCount,
        int maximumLength)
    {
        if (values is null || values.Count < minimumCount || values.Count > maximumCount)
        {
            throw new InvalidOperationException(
                $"{field} must contain between {minimumCount} and {maximumCount} items.");
        }

        foreach (var value in values)
        {
            ValidateText(value, field, 1, maximumLength);
        }
    }

    private static void ValidateIdentifierList(
        IReadOnlyList<string> values,
        string field,
        int maximumCount)
    {
        if (values is null || values.Count > maximumCount)
        {
            throw new InvalidOperationException($"{field} must not contain more than {maximumCount} items.");
        }

        foreach (var value in values)
        {
            ValidateIdentifier(value, field);
        }

        EnsureUnique(values, field);
    }

    private static void ValidateOptionalIdentifier(string? value, string field)
    {
        if (value is not null)
        {
            ValidateIdentifier(value, field);
        }
    }

    private static void ValidateIdentifier(string value, string field)
    {
        if (value is null || !SafeIdentifierRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "brand_profile_identifier_invalid",
                $"{field} must contain only ASCII letters, digits, hyphens, and underscores (maximum 128 characters).");
        }
    }

    private static void ValidateVersion(string value)
    {
        if (value is null || !VersionRegex().IsMatch(value))
        {
            throw new PptxValidationException(
                "brand_profile_version_invalid",
                "Brand Profile version must be an opaque version token of at most 64 characters.");
        }
    }

    private static void ValidateText(string value, string field, int minimumLength, int maximumLength)
    {
        if (value is null
            || value.Length < minimumLength
            || value.Length > maximumLength
            || value.Any(character => char.IsControl(character) && character is not ('\n' or '\r' or '\t'))
            || ContainsExternalLocator(value))
        {
            throw new InvalidOperationException(
                $"{field} must contain {minimumLength} to {maximumLength} characters and no URL, path, or control character.");
        }
    }

    private static bool ContainsExternalLocator(string value) =>
        value.Contains("://", StringComparison.OrdinalIgnoreCase)
        || FileSchemeRegex().IsMatch(value)
        || value.Contains("../", StringComparison.Ordinal)
        || value.Contains("..\\", StringComparison.Ordinal)
        || value.StartsWith('/')
        || value.StartsWith("\\\\", StringComparison.Ordinal)
        || AbsolutePosixPathRegex().IsMatch(value)
        || UncPathRegex().IsMatch(value)
        || WindowsPathRegex().IsMatch(value);

    private static void EnsureUnique(IEnumerable<string> values, string label)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        if (values.Any(value => !seen.Add(value)))
        {
            throw new InvalidOperationException($"Duplicate {label} IDs are not allowed.");
        }
    }

    private static void EnsureContrast(
        string foreground,
        string background,
        double minimum,
        string roles)
    {
        var ratio = ContrastRatio(foreground, background);
        if (ratio < minimum)
        {
            throw new InvalidOperationException(
                $"Brand Profile color contrast for {roles} must be at least {minimum:0.0}:1; actual ratio is {ratio:0.00}:1.");
        }
    }

    private static double ContrastRatio(string first, string second)
    {
        var firstLuminance = RelativeLuminance(first);
        var secondLuminance = RelativeLuminance(second);
        var lighter = Math.Max(firstLuminance, secondLuminance);
        var darker = Math.Min(firstLuminance, secondLuminance);
        return (lighter + 0.05) / (darker + 0.05);
    }

    private static double RelativeLuminance(string color)
    {
        var red = byte.Parse(color.AsSpan(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
        var green = byte.Parse(color.AsSpan(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
        var blue = byte.Parse(color.AsSpan(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
        return 0.2126 * Linearize(red) + 0.7152 * Linearize(green) + 0.0722 * Linearize(blue);
    }

    private static double Linearize(double component) =>
        component <= 0.04045
            ? component / 12.92
            : Math.Pow((component + 0.055) / 1.055, 2.4);

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex SafeIdentifierRegex();

    [GeneratedRegex("\\A[A-Za-z0-9][A-Za-z0-9._-]{0,63}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex VersionRegex();

    [GeneratedRegex("\\A[0-9a-f]{64}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex ContentHashRegex();

    [GeneratedRegex("\\A#[0-9A-Fa-f]{6}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex ColorRegex();

    [GeneratedRegex("(?<![A-Za-z0-9_])[A-Za-z]:[\\\\/]", RegexOptions.CultureInvariant)]
    private static partial Regex WindowsPathRegex();

    [GeneratedRegex("(?i:(?<![A-Za-z0-9_])file:)", RegexOptions.CultureInvariant)]
    private static partial Regex FileSchemeRegex();

    [GeneratedRegex("""(?<![A-Za-z0-9_])/[A-Za-z0-9._-]+(?:/[^\s"')\]}>,;:]+)*""", RegexOptions.CultureInvariant)]
    private static partial Regex AbsolutePosixPathRegex();

    [GeneratedRegex("""(?<![A-Za-z0-9_])\\\\[^\\\s"')\]}>,;:]+\\[^\s"')\]}>,;:]+""", RegexOptions.CultureInvariant)]
    private static partial Regex UncPathRegex();

    private sealed class BrandProfileManifest
    {
        [JsonPropertyName("schema_version")]
        public int SchemaVersion { get; init; }

        [JsonPropertyName("id")]
        public string Id { get; init; } = string.Empty;

        [JsonPropertyName("version")]
        public string Version { get; init; } = string.Empty;

        [JsonPropertyName("name")]
        public string Name { get; init; } = string.Empty;

        [JsonPropertyName("description")]
        public string Description { get; init; } = string.Empty;

        [JsonPropertyName("template_source")]
        public string TemplateSource { get; init; } = string.Empty;

        [JsonPropertyName("template_id")]
        public string TemplateId { get; init; } = string.Empty;

        [JsonPropertyName("color_roles")]
        public BrandColorRoles ColorRoles { get; init; } =
            new("", "", "", "", "", "", "", "", "", "", []);

        [JsonPropertyName("typography")]
        public BrandTypography Typography { get; init; } = new("", "");

        [JsonPropertyName("voice_rules")]
        public IReadOnlyList<string> VoiceRules { get; init; } = [];

        [JsonPropertyName("visual_rules")]
        public BrandVisualRuleSet VisualRules { get; init; } = new([], [], [], [], [], [], [], []);

        [JsonPropertyName("prohibited_rules")]
        public IReadOnlyList<string> ProhibitedRules { get; init; } = [];

        [JsonPropertyName("requires_confirmation_rules")]
        public IReadOnlyList<string> RequiresConfirmationRules { get; init; } = [];

        [JsonPropertyName("approved_asset_collection_ids")]
        public IReadOnlyList<string> ApprovedAssetCollectionIds { get; init; } = [];

        [JsonPropertyName("style_directions")]
        public IReadOnlyList<BrandStyleDirection> StyleDirections { get; init; } = [];

        [JsonPropertyName("layout_recipes")]
        public IReadOnlyList<BrandLayoutRecipe> LayoutRecipes { get; init; } = [];

        [JsonPropertyName("samples")]
        public IReadOnlyList<BrandSampleSummary> Samples { get; init; } = [];
    }
}
