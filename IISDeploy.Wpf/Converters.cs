using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using IISDeploy.Wpf.Models;

namespace IISDeploy.Wpf;

/// <summary>Shows just the file name of a full path.</summary>
public sealed class PathToFileNameConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is string s && !string.IsNullOrEmpty(s) ? Path.GetFileName(s) : string.Empty;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => Binding.DoNothing;
}

/// <summary>Bool -> Visibility, inverted (true =&gt; Collapsed).</summary>
public sealed class InverseBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? Visibility.Collapsed : Visibility.Visible;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => Binding.DoNothing;
}

/// <summary>Maps a log entry's severity to the appropriate foreground brush.</summary>
public sealed class SeverityToBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var key = value switch
        {
            LogSeverity.Error => "ErrorBrush",
            LogSeverity.Warning => "WarningBrush",
            LogSeverity.Success => "SuccessBrush",
            _ => "TextPrimaryBrush",
        };

        return Application.Current.TryFindResource(key) as Brush ?? Brushes.White;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => Binding.DoNothing;
}

/// <summary>
/// Progress-indicator helper: given the current 1-based step number and a target
/// step index (as ConverterParameter), returns the accent brush once that step has
/// been reached, otherwise the muted border brush.
/// </summary>
public sealed class StepReachedToBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        int current = value is int i ? i : 0;
        int target = parameter is string s && int.TryParse(s, out var p) ? p : 0;

        var key = current >= target ? "AccentBrush" : "BorderBrush";
        return Application.Current.TryFindResource(key) as Brush ?? Brushes.Gray;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => Binding.DoNothing;
}
