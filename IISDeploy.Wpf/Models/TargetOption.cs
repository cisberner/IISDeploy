using IISDeploy.Core;

namespace IISDeploy.Wpf.Models;

/// <summary>
/// One existing IIS site shown in the "update existing site" list.
/// </summary>
public sealed class TargetOption
{
    public required string Title { get; init; }
    public string? Subtitle { get; init; }
    public required SiteInfo Site { get; init; }

    // Segoe MDL2 Assets World glyph (E774) for an existing site.
    public string Glyph => char.ConvertFromUtf32(0xE774);

    public static TargetOption FromSite(SiteInfo site) => new()
    {
        Title = site.Name,
        Subtitle = string.IsNullOrWhiteSpace(site.PhysicalPath)
            ? site.State
            : site.PhysicalPath,
        Site = site,
    };
}
