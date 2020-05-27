using Playnite.SDK;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Controls;

namespace ShortcutSync
{
    public class ShortcutSyncSettings : ISettings
    {
        private readonly ShortcutSync plugin;

        public string ShortcutPath { get; set; } =
            System.Environment.ExpandEnvironmentVariables(Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.StartMenu), "PlayniteGames"));

        public bool InstalledOnly { get; set; } = true;
        public bool UsePlayAction { get; set; } = false;
        public bool UpdateOnStartup { get; set; } = false;
        public bool ForceUpdate { get; set; } = false;
        public Dictionary<string, bool> SourceOptions { get; set; } = new Dictionary<string, bool>();

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
                ShortcutPath = savedSettings.ShortcutPath;
                SourceOptions = savedSettings.SourceOptions;
                UpdateOnStartup = savedSettings.UpdateOnStartup;
                UsePlayAction = savedSettings.UsePlayAction;
            }
        }

        public void BeginEdit()
        {
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
            ShortcutPath = plugin.PlayniteApi.Dialogs.SelectFolder();
            ((TextBox)plugin.settingsView.FindName("PathTextBox")).Text = ShortcutPath;
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
            errors = new List<string>();
            try
            {
                Directory.CreateDirectory(ShortcutPath);
            }
            catch (Exception ex)
            {
                errors.Add(ex.Message);
                return false;
            }
            return true;
        }
    }
}