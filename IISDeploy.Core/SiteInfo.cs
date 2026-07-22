namespace IISDeploy.Core;

/// <summary>
/// Lightweight, disconnected description of an installed IIS site,
/// suitable for presenting a selection list in any UI.
/// </summary>
public sealed class SiteInfo
{
    public required string Name { get; init; }
    public string? PhysicalPath { get; init; }
    public string? State { get; init; }

    /// <summary>
    /// A browseable URL built from the site's first http/https binding (https preferred),
    /// or null when the site has no such binding. Used to open the site in a browser.
    /// </summary>
    public string? Url { get; init; }
}
