using System.Windows;

namespace IISDeploy.Wpf.Views;

/// <summary>
/// Modal editor for a single configuration file (appsettings.json or web.config).
/// Resizable, with a large monospace text area for long files.
/// </summary>
public partial class FileEditorWindow : Window
{
    public FileEditorWindow(string fileName, string content)
    {
        InitializeComponent();
        Title = $"Edit {fileName}";
        HeaderText.Text = $"Edit {fileName}";
        EditorBox.Text = content ?? string.Empty;
    }

    public string EditedText => EditorBox.Text;

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
