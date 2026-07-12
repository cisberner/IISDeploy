using System.Windows;

namespace IISDeploy.Wpf.Views;

/// <summary>
/// Modal editor for the new site's appsettings.json and web.config contents.
/// </summary>
public partial class ConfigFilesWindow : Window
{
    public ConfigFilesWindow(string appSettings, string webConfig)
    {
        InitializeComponent();
        AppSettingsBox.Text = appSettings ?? string.Empty;
        WebConfigBox.Text = webConfig ?? string.Empty;
    }

    public string AppSettingsText => AppSettingsBox.Text;
    public string WebConfigText => WebConfigBox.Text;

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
