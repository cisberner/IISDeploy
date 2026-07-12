using System.Collections.Specialized;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using IISDeploy.Wpf.ViewModels;

namespace IISDeploy.Wpf;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel = new();

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;

        // Keep the log view pinned to the newest line as entries stream in.
        _viewModel.LogEntries.CollectionChanged += OnLogEntriesChanged;
    }

    private void OnLogEntriesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.Action == NotifyCollectionChangedAction.Add)
            Dispatcher.BeginInvoke(new Action(() => LogScrollViewer.ScrollToEnd()));
    }

    private void InstallModeCard_Checked(object sender, RoutedEventArgs e)
    {
        // Move focus to the site-name box when the user picks "Install new site".
        // Deferred to Input priority so the form is laid out and visible first.
        Dispatcher.BeginInvoke(new Action(() =>
        {
            NewSiteNameBox.Focus();
            Keyboard.Focus(NewSiteNameBox);
        }), DispatcherPriority.Input);
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        => WindowState = WindowState.Minimized;

    private void CloseButton_Click(object sender, RoutedEventArgs e)
        => Close();
}
