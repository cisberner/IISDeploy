using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace IISDeploy.Wpf;

public partial class StepPill : UserControl
{
    public StepPill()
    {
        InitializeComponent();
        Loaded += (_, _) => UpdateVisualState();
    }

    public static readonly DependencyProperty NumberProperty = DependencyProperty.Register(
        nameof(Number), typeof(int), typeof(StepPill),
        new PropertyMetadata(1, OnStateChanged));

    public static readonly DependencyProperty CaptionProperty = DependencyProperty.Register(
        nameof(Caption), typeof(string), typeof(StepPill),
        new PropertyMetadata("Step", OnStateChanged));

    public static readonly DependencyProperty StepNumberProperty = DependencyProperty.Register(
        nameof(StepNumber), typeof(int), typeof(StepPill),
        new PropertyMetadata(1, OnStateChanged));

    public int Number
    {
        get => (int)GetValue(NumberProperty);
        set => SetValue(NumberProperty, value);
    }

    public string Caption
    {
        get => (string)GetValue(CaptionProperty);
        set => SetValue(CaptionProperty, value);
    }

    /// <summary>The wizard's current 1-based step number.</summary>
    public int StepNumber
    {
        get => (int)GetValue(StepNumberProperty);
        set => SetValue(StepNumberProperty, value);
    }

    private static void OnStateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        => ((StepPill)d).UpdateVisualState();

    private void UpdateVisualState()
    {
        if (NumberText == null) // template not applied yet
            return;

        NumberText.Text = Number.ToString();
        CaptionText.Text = Caption;

        bool done = StepNumber > Number;
        bool active = StepNumber == Number;

        var accent = TryBrush("AccentGradientBrush") ?? TryBrush("AccentBrush");
        var surface = TryBrush("SurfaceAltBrush");
        var border = TryBrush("BorderBrush");
        var textPrimary = TryBrush("TextPrimaryBrush");
        var textSecondary = TryBrush("TextSecondaryBrush");

        if (done || active)
        {
            Circle.Background = accent;
            Circle.BorderThickness = new Thickness(0);
            CaptionText.Foreground = textPrimary;
        }
        else
        {
            Circle.Background = surface;
            Circle.BorderBrush = border;
            Circle.BorderThickness = new Thickness(1);
            CaptionText.Foreground = textSecondary;
        }

        CheckText.Visibility = done ? Visibility.Visible : Visibility.Collapsed;
        NumberText.Visibility = done ? Visibility.Collapsed : Visibility.Visible;
        NumberText.Foreground = active ? Brushes.White : textSecondary;
    }

    private Brush? TryBrush(string key) => Application.Current?.TryFindResource(key) as Brush;
}
