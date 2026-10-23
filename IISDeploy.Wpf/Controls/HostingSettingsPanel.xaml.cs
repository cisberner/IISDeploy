using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace IISDeploy.Wpf;

public partial class HostingSettingsPanel : UserControl
{
    public HostingSettingsPanel()
    {
        InitializeComponent();
    }

    // Keep the idle time-out numeric: block any typed character that is not a digit...
    private void NumberBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        e.Handled = !e.Text.All(char.IsDigit);
    }

    // ...and reject a paste unless it is entirely digits.
    private void NumberBox_Pasting(object sender, DataObjectPastingEventArgs e)
    {
        if (e.DataObject.GetDataPresent(DataFormats.UnicodeText)
            && e.DataObject.GetData(DataFormats.UnicodeText) is string text
            && text.Length > 0 && text.All(char.IsDigit))
        {
            return;
        }

        e.CancelCommand();
    }
}
