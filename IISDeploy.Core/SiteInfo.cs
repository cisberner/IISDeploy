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
}
