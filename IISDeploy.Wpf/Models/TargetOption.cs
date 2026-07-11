using IISDeploy.Core;

namespace IISDeploy.Wpf.Models;

public enum TargetKind
{
    ExistingSite,
    CreateNew,
}

/// <summary>
/// One selectable deployment target on the Configure page: either an existing IIS
/// site or the "create a new site" option.
/// </summary>
public sealed class TargetOption
{
    public required TargetKind Kind { get; init; }
    public required string Title { get; init; }
    public string? Subtitle { get; init; }
    public SiteInfo? Site { get; init; }

    public bool IsCreateNew => Kind == TargetKind.CreateNew;

    // Segoe MDL2 Assets glyphs: Add (E710) for the new-site option, World (E774)
    // for existing sites.
    public string Glyph => IsCreateNew ? char.ConvertFromUtf32(0xE710) : char.ConvertFromUtf32(0xE774);

    public static TargetOption FromSite(SiteInfo site) => new()
    {
        Kind = TargetKind.ExistingSite,
        Title = site.Name,
        Subtitle = string.IsNullOrWhiteSpace(site.PhysicalPath)
            ? site.State
            : site.PhysicalPath,
        Site = site,
    };

    public static TargetOption CreateNewOption() => new()
    {
        Kind = TargetKind.CreateNew,
        Title = "Create a new site",
        Subtitle = "Set up a new IIS site with an HTTPS binding and self-signed certificate",
    };
}
