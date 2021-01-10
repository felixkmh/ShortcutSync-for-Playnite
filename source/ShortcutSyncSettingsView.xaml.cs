using System.Diagnostics;
using System.Windows.Controls;


namespace ShortcutSync
{
    public partial class ShortcutSyncSettingsView : UserControl
    {
        public ShortcutSyncSettingsView()
        {
            InitializeComponent();
        }

        private void URL_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}