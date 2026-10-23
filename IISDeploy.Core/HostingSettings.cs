using Microsoft.Web.Administration;

namespace IISDeploy.Core;

/// <summary>
/// The IIS settings that keep a site's application warm: the application pool's Start Mode
/// and Idle Time-out, and the site's Preload Enabled. The defaults start the worker process
/// with IIS instead of with the first request, never shut it down when idle, and warm the
/// application up as soon as it starts. Applied to a new site and on every update of an
/// existing one.
/// </summary>
public sealed class HostingSettings
{
    /// <summary>The largest Idle Time-out IIS accepts: 30 days, in minutes.</summary>
    public const int MaxIdleTimeoutMinutes = 43200;

    /// <summary>
    /// IIS: App Pool > Advanced Settings > Start Mode. AlwaysRunning starts the worker process
    /// with IIS; OnDemand waits for the first request.
    /// </summary>
    public StartMode StartMode { get; init; } = StartMode.AlwaysRunning;

    /// <summary>
    /// IIS: App Pool > Advanced Settings > Idle Time-out (minutes). 0 never shuts an idle
    /// worker process down.
    /// </summary>
    public int IdleTimeoutMinutes { get; init; }

    /// <summary>
    /// IIS: Site > Advanced Settings > Preload Enabled. Warms the application up as soon as
    /// the worker process starts; needs the IIS feature Application Initialization.
    /// </summary>
    public bool PreloadEnabled { get; init; } = true;

    /// <summary>
    /// Returns why these settings cannot be applied, or null when they can.
    /// </summary>
    public string? Validate()
    {
        if (StartMode != StartMode.AlwaysRunning && StartMode != StartMode.OnDemand)
            return $"Invalid start mode '{StartMode}'.";

        if (IdleTimeoutMinutes < 0 || IdleTimeoutMinutes > MaxIdleTimeoutMinutes)
            return $"Invalid idle time-out '{IdleTimeoutMinutes}'. Valid range is 0-{MaxIdleTimeoutMinutes} minutes.";

        return null;
    }
}
