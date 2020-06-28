using Playnite.SDK;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Controls;
using System.Windows.Documents;

namespace ShortcutSync
{
    public class ShortcutSyncSettings : ISettings
    {
        public delegate void OnPathChangedAction(string newPath);
        public event OnPathChangedAction OnPathChanged;
        private readonly ShortcutSync plugin;

        public string ShortcutPath { get; set; } =
            System.Environment.ExpandEnvironmentVariables(Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.StartMenu), "PlayniteGames"));

        public bool InstalledOnly { get; set; } = true;
        public bool UsePlayAction { get; set; } = false;
        public bool UpdateOnStartup { get; set; } = false;
        public bool ForceUpdate { get; set; } = false;
        public bool ExcludeHidden { get; set; } = false;
        public Dictionary<string, bool> SourceOptions { get; set; } = new Dictionary<string, bool>() { { "Undefined", false } };

        private string tempShortcutPath;

        // Parameterless constructor must exist if you want to use LoadPluginSettings method.
        public ShortcutSyncSettings()
        {

        }

        public ShortcutSyncSettings(ShortcutSync plugin)
        {
            // Injecting your plugin instance is required for Save/Load method because Playnite saves data to a location based on what plugin requested the operation.
            this.plugin = plugin;

            // Load saved settings.
            var savedSettings = plugin.LoadPluginSettings<ShortcutSyncSettings>();

            // LoadPluginSettings returns null if not saved data is available.
            if (savedSettings != null)
            {
                InstalledOnly = savedSettings.InstalledOnly;
                ForceUpdate = savedSettings.ForceUpdate;
                if (savedSettings.ShortcutPath != null) 
                    ShortcutPath = savedSettings.ShortcutPath;
                if (savedSettings.SourceOptions != null)
                    SourceOptions = savedSettings.SourceOptions;
                UpdateOnStartup = savedSettings.UpdateOnStartup;
                UsePlayAction = savedSettings.UsePlayAction;
                ExcludeHidden = savedSettings.ExcludeHidden;
            }
        }

        public void BeginEdit()
        {
            var updateLabel = (Run)plugin.settingsView.FindName("UpdateLabel");
            var urlHyperlink = (Hyperlink)plugin.settingsView.FindName("URL");
            var urlLabel = (TextBlock)plugin.settingsView.FindName("URLLabel");
            var updateTextBlock = (TextBlock)plugin.settingsView.FindName("UpdateTextBlock");
            updateLabel.Text = "ShortcutSync Up-To-Date.";
            urlLabel.Text = "";
            var updateAvailableResult = plugin.UpdateAvailable().Result; 
            if (updateAvailableResult.Available)
            {
                updateLabel.Text = $"New version {updateAvailableResult.LatestVersion} available at:\n";
                urlHyperlink.NavigateUri = new Uri(updateAvailableResult.Url);
                urlLabel.Text = updateAvailableResult.Url;
            }


            var bt = (Button)plugin.settingsView.FindName("SelectFolderButton");
            bt.Click += Bt_Click;
            // Get all available source names
            foreach (var src in plugin.PlayniteApi.Database.Sources)
            {
                // Add new sources and disable them by default
                if (!SourceOptions.ContainsKey(src.Name))
                {
                    SourceOptions.Add(src.Name, false);
                }
            }
            if (!SourceOptions.ContainsKey(Constants.UNDEFINEDSOURCE))
            {
                SourceOptions.Add(Constants.UNDEFINEDSOURCE, false);
            }
            // Set view up
            var container = (StackPanel)plugin.settingsView.FindName("SourceSettingsStack");
            container.Children.Clear();
            foreach (var srcOpt in SourceOptions)
            {
                // Add checkboxes and set some properties
                var checkBox = new CheckBox();
                checkBox.Content = srcOpt.Key;
                checkBox.IsChecked = srcOpt.Value;
                checkBox.MinWidth = 150;
                checkBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                container.Children.Add(checkBox);
            }
        }

        // Callback for the Select Folder Button
        private void Bt_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string path = plugin.PlayniteApi.Dialogs.SelectFolder();
            if (!string.IsNullOrEmpty(path))
            {
                ((TextBox)plugin.settingsView.FindName("PathTextBox")).Text = path;
            }
        }

        public void CancelEdit()
        {
            // Code executed when user decides to cancel any changes made since BeginEdit was called.
            // This method should revert any changes made to Option1 and Option2.
        }

        public void EndEdit()
        {
            var bt = (Button)plugin.settingsView.FindName("SelectFolderButton");
            bt.Click -= Bt_Click;
            var container = (StackPanel)plugin.settingsView.FindName("SourceSettingsStack");
            foreach (CheckBox checkBox in container.Children)
            {
                SourceOptions[checkBox.Content.ToString()] = (bool)checkBox.IsChecked;
            }
            // Code executed when user decides to confirm changes made since BeginEdit was called.
            plugin.SavePluginSettings(this);
        }

        public bool VerifySettings(out List<string> errors)
        {
            // Code execute when user decides to confirm changes made since BeginEdit was called.
            // Executed before EndEdit is called and EndEdit is not called if false is returned.
            // List of errors is presented to user if verification fails.
            tempShortcutPath = ((TextBox)plugin.settingsView.FindName("PathTextBox")).Text;
            errors = new List<string>();
            try
            {
                Directory.CreateDirectory(tempShortcutPath);
            }
            catch (Exception ex)
            {
                errors.Add(ex.Message);
                return false;
            }
            // Check whether user has write persmission to the selected folder
            try
            {
                var ac = Directory.GetAccessControl(tempShortcutPath);
            }
            catch (UnauthorizedAccessException)
            {
                errors.Add("The selected folder cannot be written to. Please select another folder.");
                return false;
            }
            if (ShortcutPath != tempShortcutPath)
                OnPathChanged?.Invoke(tempShortcutPath);
            ShortcutPath = tempShortcutPath;
            return true;
        }
    }
}