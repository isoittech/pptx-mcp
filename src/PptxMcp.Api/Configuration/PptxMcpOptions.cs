namespace PptxMcp.Configuration;

public sealed class PptxMcpOptions
{
    public const string SectionName = "PptxMcp";

    public string SharedSecret { get; init; } = string.Empty;

    public string SigningKey { get; init; } = string.Empty;

    public string PublicBaseUrl { get; init; } = "http://localhost:8080";

    public string StorageRoot { get; init; } = "/data/pptx-mcp";

    public string LibreChatUploadsRoot { get; init; } = "/data/librechat-uploads";

    public string LibreChatImagesRoot { get; init; } = "/data/librechat-images";

    public string TemplatesRoot { get; init; } = "/data/pptx-templates";

    public string DefaultTemplateId { get; init; } = string.Empty;

    public int DefaultTemplateCoverSampleSlideNumber { get; init; }

    public int DefaultTemplateBodySampleSlideNumber { get; init; }

    public string BrandProfilesRoot { get; init; } = "/data/pptx-brand-profiles";

    public bool RequireDesignBrief { get; init; }

    public int DesignBriefLifetimeMinutes { get; init; } = 60;

    public string FirstAssistantNotice { get; init; } = string.Empty;

    public string VisualRendererPath { get; init; } = "/app/visual-renderer/index.mjs";

    public string ImageSanitizerPath { get; init; } = "/app/visual-renderer/sanitize-image.mjs";

    public long MaxFileBytes { get; init; } = 30L * 1024 * 1024;

    public long MaxImageFileBytes { get; init; } = 12L * 1024 * 1024;

    public int MaxImagePixels { get; init; } = 20_000_000;

    public int MaxImageDimension { get; init; } = 2_560;

    public int MaxSlides { get; init; } = 50;

    public int MaxConcurrentJobs { get; init; } = 3;

    public int MaxQueueDepth { get; init; } = 12;

    public int JobTimeoutMinutes { get; init; } = 10;

    public int RetentionDays { get; init; } = 7;

    public int RetentionHoursAfterDownload { get; init; } = 24;

    public int ArtifactUrlMinutes { get; init; } = 15;

    public int MaxZipEntries { get; init; } = 5_000;

    public long MaxUncompressedBytes { get; init; } = 300L * 1024 * 1024;

    public int MaxCompressionRatio { get; init; } = 250;

    public void Validate(bool requireSecrets)
    {
        if (requireSecrets && SharedSecret.Length < 24)
        {
            throw new InvalidOperationException("PptxMcp:SharedSecret must contain at least 24 characters.");
        }

        if (requireSecrets && SigningKey.Length < 32)
        {
            throw new InvalidOperationException("PptxMcp:SigningKey must contain at least 32 characters.");
        }

        if (!Uri.TryCreate(PublicBaseUrl, UriKind.Absolute, out _))
        {
            throw new InvalidOperationException("PptxMcp:PublicBaseUrl must be an absolute URL.");
        }

        if (MaxFileBytes <= 0 || MaxSlides is <= 0 or > 50 || MaxConcurrentJobs is <= 0 or > 3)
        {
            throw new InvalidOperationException("Configured resource limits are outside the supported bounds.");
        }

        if (JobTimeoutMinutes is <= 0 or > 10)
        {
            throw new InvalidOperationException("PptxMcp:JobTimeoutMinutes must be between 1 and 10.");
        }

        if (!Path.IsPathFullyQualified(VisualRendererPath))
        {
            throw new InvalidOperationException("PptxMcp:VisualRendererPath must be an absolute path.");
        }

        if (!Path.IsPathFullyQualified(ImageSanitizerPath))
        {
            throw new InvalidOperationException("PptxMcp:ImageSanitizerPath must be an absolute path.");
        }

        if (!Path.IsPathFullyQualified(LibreChatUploadsRoot)
            || !Path.IsPathFullyQualified(LibreChatImagesRoot))
        {
            throw new InvalidOperationException(
                "PptxMcp LibreChat upload roots must be absolute paths.");
        }

        if (MaxImageFileBytes is <= 0 or > 30L * 1024 * 1024
            || MaxImagePixels is <= 0 or > 40_000_000
            || MaxImageDimension is < 320 or > 4_096)
        {
            throw new InvalidOperationException("Configured image resource limits are outside the supported bounds.");
        }

        if (!Path.IsPathFullyQualified(TemplatesRoot))
        {
            throw new InvalidOperationException("PptxMcp:TemplatesRoot must be an absolute path.");
        }

        if (!Path.IsPathFullyQualified(BrandProfilesRoot))
        {
            throw new InvalidOperationException("PptxMcp:BrandProfilesRoot must be an absolute path.");
        }

        if (DesignBriefLifetimeMinutes is < 5 or > 120)
        {
            throw new InvalidOperationException(
                "PptxMcp:DesignBriefLifetimeMinutes must be between 5 and 120.");
        }

        if (DefaultTemplateId.Length > 128
            || DefaultTemplateId.Any(character => !char.IsAsciiLetterOrDigit(character) && character is not '-' and not '_'))
        {
            throw new InvalidOperationException(
                "PptxMcp:DefaultTemplateId may contain only ASCII letters, digits, hyphens, and underscores (maximum 128 characters).");
        }

        var hasCoverSample = DefaultTemplateCoverSampleSlideNumber > 0;
        var hasBodySample = DefaultTemplateBodySampleSlideNumber > 0;
        if (hasCoverSample != hasBodySample
            || DefaultTemplateCoverSampleSlideNumber is < 0 or > 50
            || DefaultTemplateBodySampleSlideNumber is < 0 or > 50)
        {
            throw new InvalidOperationException(
                "Default template cover/body sample slide numbers must both be unset or both be between 1 and 50.");
        }

        if (FirstAssistantNotice.Length > 1_000)
        {
            throw new InvalidOperationException("PptxMcp:FirstAssistantNotice must not exceed 1000 characters.");
        }
    }
}
