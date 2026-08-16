using PptxMcp.Design;
using PptxMcp.Domain;
using PptxMcp.Storage;

namespace PptxMcp.Tests;

public sealed class BrandProfileCatalogTests
{
    [Fact]
    public void ReturnsCompactListThenFilteredRecipesAndSamples()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);

            var compact = catalog.Query();
            var filtered = catalog.Query(
                profileId: "general",
                purpose: "kpi",
                density: "balanced",
                styleDirectionId: "standard");

            Assert.Equal("available", compact.Status);
            var compactProfile = Assert.Single(compact.Profiles);
            Assert.Null(compactProfile.Detail);
            Assert.Empty(compactProfile.Recipes);
            Assert.Empty(compactProfile.Samples);
            Assert.Matches("^[0-9a-f]{64}$", compactProfile.Summary.ContentHash);
            var direction = Assert.Single(compactProfile.Summary.StyleDirections);
            Assert.Equal("standard", direction.Id);
            Assert.Equal("executive", direction.DesignStyle);
            Assert.Contains("exactly once more", compact.Instruction, StringComparison.Ordinal);

            var selected = Assert.Single(filtered.Profiles);
            Assert.NotNull(selected.Detail);
            Assert.Equal("kpi-balanced", Assert.Single(selected.Recipes).Id);
            Assert.Equal("sample-kpi-medium", Assert.Single(selected.Samples).Id);
            Assert.Equal(compactProfile.Summary.ContentHash, selected.Summary.ContentHash);
            Assert.Contains("Do not call pptx_get_design_catalog again", filtered.Instruction, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RequiresProfileIdBeforeApplyingRecipeFilters()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);

            var error = Assert.Throws<PptxValidationException>(() =>
                catalog.Query(purpose: "kpi", density: "balanced"));

            Assert.Equal("design_catalog_profile_required", error.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void KeepsImmutableInProcessSnapshotAfterExternalManifestChanges()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root, name: "Original profile");

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);
            var original = catalog.Query("general").Profiles.Single().Summary;

            BrandProfileTestFactory.WriteProfile(root, name: "Changed profile");
            var repeated = catalog.Query("general").Profiles.Single().Summary;

            Assert.Equal("Original profile", repeated.Name);
            Assert.Equal(original.ContentHash, repeated.ContentHash);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ExactProfileReferenceRejectsStaleContentHash()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);
            var summary = catalog.Query("general").Profiles.Single().Summary;
            var stale = new BrandProfileReference(
                summary.Id,
                summary.Version,
                new string('0', 64));

            var error = Assert.Throws<PptxValidationException>(() => catalog.GetSnapshot(stale));

            Assert.Equal("brand_profile_version_mismatch", error.Code);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsUrlsAndPathsInExternalProfileText()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(
            root,
            description: "Read internal guidance at https://internal.invalid/brand.");

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);

            var error = Assert.Throws<InvalidOperationException>(catalog.EnsureReady);

            Assert.Contains("no URL, path", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData("Use the asset at (/mnt/brand/approved.png).")]
    [InlineData("Use the asset at /logo.png.")]
    [InlineData("Do not reveal /secret.")]
    [InlineData("Use the asset at \\\\server\\share\\approved.png.")]
    [InlineData("Use the asset at file:/mnt/brand/approved.png.")]
    public void RejectsEmbeddedAbsoluteFileLocatorsInExternalProfileText(string description)
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root, description: description);

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);

            var error = Assert.Throws<InvalidOperationException>(catalog.EnsureReady);

            Assert.Contains("no URL, path", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RequiredBriefFailsClosedWhenCatalogRootIsMissing()
    {
        var missingRoot = Path.Combine(Path.GetTempPath(), $"missing-brand-profiles-{Guid.NewGuid():N}");
        var catalog = BrandProfileTestFactory.CreateCatalog(missingRoot, requireDesignBrief: true);

        var error = Assert.Throws<InvalidOperationException>(catalog.EnsureReady);

        Assert.Contains("does not exist", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsUnreadableTextAndSurfaceRoleContrast()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root);
        var manifestPath = Path.Combine(root, "general", "brand-profile.json");
        var manifest = File.ReadAllText(manifestPath);
        File.WriteAllText(
            manifestPath,
            manifest.Replace("#17212B", "#FFFFFF", StringComparison.Ordinal));

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);

            var error = Assert.Throws<InvalidOperationException>(catalog.EnsureReady);

            Assert.Contains("text/background", error.Message, StringComparison.Ordinal);
            Assert.Contains("4.5:1", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsRecipeDensityOutsideReferencedStyleDirection()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root, supportedDensities: ["balanced"]);

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);

            var error = Assert.Throws<InvalidOperationException>(catalog.EnsureReady);

            Assert.Contains("supported_densities", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsMetricsSpotlightRecipeWithoutExactlyThreeItems()
    {
        var root = BrandProfileTestFactory.CreateRoot();
        BrandProfileTestFactory.WriteProfile(root, kpiItemCount: 4);

        try
        {
            var catalog = BrandProfileTestFactory.CreateCatalog(root);

            var error = Assert.Throws<InvalidOperationException>(catalog.EnsureReady);

            Assert.Contains("item_count", error.Message, StringComparison.Ordinal);
            Assert.Contains("exactly 3", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }
}
