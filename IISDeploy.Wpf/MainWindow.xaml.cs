using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
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

        // PasswordBox.Password is not bindable; mirror the VM value when it is reset
        // (e.g. Start over) so a stale password never lingers in the box.
        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.AppPoolPassword)
            && AppPoolPasswordBox.Password != _viewModel.AppPoolPassword)
        {
            AppPoolPasswordBox.Password = _viewModel.AppPoolPassword;
        }
    }

    private void AppPoolPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        _viewModel.AppPoolPassword = ((PasswordBox)sender).Password;
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

    private void UseCustomIdentity_Checked(object sender, RoutedEventArgs e)
    {
        // Move focus to the user-name box when the custom-account option is enabled.
        // Deferred to Input priority so the field is laid out and visible first.
        Dispatcher.BeginInvoke(new Action(() =>
        {
            AppPoolUserNameBox.Focus();
            Keyboard.Focus(AppPoolUserNameBox);
        }), DispatcherPriority.Input);
    }

    // Keep the port box numeric: block any typed character that is not a digit...
    private void PortBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        e.Handled = !e.Text.All(char.IsDigit);
    }

    // ...and reject a paste unless it is entirely digits.
    private void PortBox_Pasting(object sender, DataObjectPastingEventArgs e)
    {
        if (e.DataObject.GetDataPresent(DataFormats.UnicodeText)
            && e.DataObject.GetData(DataFormats.UnicodeText) is string text
            && text.Length > 0 && text.All(char.IsDigit))
        {
            return;
        }

        e.CancelCommand();
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        => WindowState = WindowState.Minimized;

    private void CloseButton_Click(object sender, RoutedEventArgs e)
        => Close();
}
