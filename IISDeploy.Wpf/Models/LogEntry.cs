namespace IISDeploy.Wpf.Models;

public enum LogSeverity
{
    Info,
    Success,
    Warning,
    Error,
}

/// <summary>
/// A single line of deployment output. The severity is inferred from the text so
/// the UI can colour errors, warnings and success messages differently, while the
/// message itself stays identical to what the CLI prints.
/// </summary>
public sealed class LogEntry
{
    public string Text { get; }
    public LogSeverity Severity { get; }
    public DateTime Timestamp { get; } = DateTime.Now;

    public string TimestampText => Timestamp.ToString("HH:mm:ss");

    public LogEntry(string text)
    {
        Text = text ?? string.Empty;
        Severity = Classify(Text);
    }

    private static LogSeverity Classify(string text)
    {
        if (text.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase))
            return LogSeverity.Error;
        if (text.StartsWith("WARNING", StringComparison.OrdinalIgnoreCase))
            return LogSeverity.Warning;

        // Milestone / completion messages emitted by DeploymentService.
        if (text.Contains("complete", StringComparison.OrdinalIgnoreCase)
            || text.Contains("Backup created", StringComparison.OrdinalIgnoreCase)
            || text.Contains("started", StringComparison.OrdinalIgnoreCase)
            || text.Contains("bound to HTTPS", StringComparison.OrdinalIgnoreCase)
            || text.Contains("registered with HTTP.sys", StringComparison.OrdinalIgnoreCase)
            || text.Contains("extracted", StringComparison.OrdinalIgnoreCase)
            || text.Equals("Done.", StringComparison.OrdinalIgnoreCase))
            return LogSeverity.Success;

        return LogSeverity.Info;
    }
}
