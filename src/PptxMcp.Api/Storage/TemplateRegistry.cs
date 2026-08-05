using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using PptxMcp.Configuration;

namespace PptxMcp.Storage;

public sealed partial class TemplateRegistry(
    IOptions<PptxMcpOptions> options,
    PptxPackageGuard packageGuard)
{
    private readonly PptxMcpOptions options = options.Value;

    public bool HasDefault => !string.IsNullOrWhiteSpace(options.DefaultTemplateId);

    public string? DefaultTemplateId => HasDefault ? options.DefaultTemplateId : null;

    public async Task<ValidatedInput?> TryResolveDefaultAsync(CancellationToken cancellationToken)
    {
        if (!HasDefault)
        {
            return null;
        }

        return await ResolveDefaultAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<ValidatedInput> ResolveDefaultAsync(CancellationToken cancellationToken)
    {
        if (!HasDefault)
        {
            throw new PptxValidationException(
                "default_template_not_configured",
                "No default PowerPoint template is configured for this deployment.");
        }

        var templateId = options.DefaultTemplateId;
        if (!SafeIdentifier().IsMatch(templateId))
        {
            throw new PptxValidationException(
                "invalid_default_template_id",
                "The configured default PowerPoint template identifier is invalid.");
        }

        var templatesRoot = Path.GetFullPath(options.TemplatesRoot);
        var templatePath = Path.GetFullPath(Path.Combine(templatesRoot, $"{templateId}.pptx"));
        var rootPrefix = templatesRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!templatePath.StartsWith(rootPrefix, StringComparison.Ordinal))
        {
            throw new PptxValidationException(
                "invalid_default_template_path",
                "The configured default PowerPoint template path is outside the template directory.");
        }

        var file = new FileInfo(templatePath);
        if (!file.Exists || file.LinkTarget is not null)
        {
            throw new PptxValidationException(
                "default_template_not_found",
                "The configured default PowerPoint template was not found.");
        }

        var input = await packageGuard.ValidateAsync(templatePath, cancellationToken).ConfigureAwait(false);
        return input with { FileId = templateId };
    }

    [GeneratedRegex("\\A[A-Za-z0-9_-]{1,128}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex SafeIdentifier();
}
