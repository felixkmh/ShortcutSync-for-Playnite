using Playnite.SDK;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;

namespace ShortcutSync
{
    public class ShortcutSyncSettings : ISettings
    {
        public delegate void OnPathChangedAction(string oldPath, string newPath);
        public event OnPathChangedAction OnPathChanged;

        public delegate void OnSettingsChangedAction();
        public event OnSettingsChangedAction OnSettingsChanged;
        private readonly ShortcutSync plugin;

        public string ShortcutPath { get; set; } =
            System.Environment.ExpandEnvironmentVariables(Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.StartMenu), "PlayniteGames"));

        public bool InstalledOnly { get; set; } = true;
        public bool UpdateOnStartup { get; set; } = false;
        public bool ForceUpdate { get; set; } = false;
        public bool ExcludeHidden { get; set; } = false;
        public bool SeparateFolders { get; set; } = false;
        public bool FadeBottom { get; set; } = false;
        public Dictionary<string, bool> SourceOptions { get; set; } = new Dictionary<string, bool>() { { "Undefined", false } };
        public Dictionary<string, bool> EnabledPlayActions { get; set; } = new Dictionary<string, bool>() { { "Undefined", false } };
        public HashSet<Guid> ManuallyCreatedShortcuts { get; set; } = new HashSet<Guid>();
        public HashSet<Guid> ExcludedGames { get; set; } = new HashSet<Guid>();


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
                var type = savedSettings.GetType();
                foreach(var prop in type.GetProperties())
                {
                    prop.SetValue(this, prop.GetValue(savedSettings));
                }
            }
        }

        public void BeginEdit()
        {
            plugin.settingsView.UpdateTextBlock.Visibility = System.Windows.Visibility.Collapsed;
            plugin.UpdateAvailable().ContinueWith((task) =>
            {
                var updateAvailable = task.Result.Available;
                var url = task.Result.Url;
                var latestVersion = task.Result.LatestVersion;
                plugin.settingsView.Dispatcher.Invoke(() =>
                {
                    if (updateAvailable)
                    {
                        plugin.settingsView.UpdateTextBlock.Visibility = System.Windows.Visibility.Visible;
                        plugin.settingsView.UpdateLabel.Text = $"New version {latestVersion} available at:\n";
                        plugin.settingsView.URL.NavigateUri = new Uri(url);
                        plugin.settingsView.URLLabel.Text = url;
                    }
                });
            });

            var bt = plugin.settingsView.SelectFolderButton;
            bt.Click += Bt_Click;
            // Get all available source names
            foreach (var src in plugin.PlayniteApi.Database.Sources)
            {
                // Add new sources and disable them by default
                if (!SourceOptions.ContainsKey(src.Name))
                {
                    SourceOptions.Add(src.Name, false);
                }
                if (!EnabledPlayActions.ContainsKey(src.Name))
                {
                    EnabledPlayActions.Add(src.Name, false);
                }
            }
            if (!SourceOptions.ContainsKey(Constants.UNDEFINEDSOURCE))
            {
                SourceOptions.Add(Constants.UNDEFINEDSOURCE, false);
            }
            if (!EnabledPlayActions.ContainsKey(Constants.UNDEFINEDSOURCE))
            {
                EnabledPlayActions.Add(Constants.UNDEFINEDSOURCE, false);
            }
            // Set up view
            var container = plugin.settingsView.SourceNamesStack;
            container.Children.Clear();
            foreach (var srcOpt in SourceOptions)
            {
                // Add label and set some properties
                var label = new Label();
                label.Content = srcOpt.Key;
                label.MinWidth = 50;
                label.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                label.Margin = new System.Windows.Thickness { Bottom = 2, Left = 5, Right = 5, Top = 2 };
                label.Height = 20;
                label.Tag = srcOpt.Key;
                container.Children.Add(label);
            }
            container = plugin.settingsView.SourceSettingsStack;
            container.Children.Clear();
            foreach (var srcOpt in SourceOptions)
            {
                // Add checkboxes and set some properties
                var checkBox = new CheckBox();
                checkBox.Content = null;
                checkBox.IsChecked = srcOpt.Value;
                checkBox.MinWidth = 50;
                checkBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                checkBox.Margin = new System.Windows.Thickness { Bottom = 2, Left = 5, Right = 5, Top = 2 };
                checkBox.Height = 20;
                checkBox.Tag = srcOpt.Key;
                checkBox.ToolTip = $"If enabled, shortcuts to games from {srcOpt.Key} are created and maintained on update.";
                container.Children.Add(checkBox);
            }
            plugin.settingsView.PlayActionSettingsStack.Children.Clear();
            foreach (var playActionOpt in EnabledPlayActions)
            {
                // Add checkboxes and set some properties
                var checkBox = new CheckBox();
                checkBox.Content = null;
                checkBox.IsChecked = playActionOpt.Value;
                checkBox.MinWidth = 50;
                checkBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                checkBox.Margin = new System.Windows.Thickness { Bottom = 2, Left = 5, Right = 5, Top = 2 };
                checkBox.Height = 20;
                checkBox.Tag = playActionOpt.Key;
                checkBox.ToolTip = $"If enabled, shortuts try to start games from {playActionOpt.Key} with their native launcher or without any launcher (bypassing Playnite), depending on the existing PlayAction of that game.";
                plugin.settingsView.PlayActionSettingsStack.Children.Add(checkBox);
            }
            plugin.settingsView.ManuallyCreatedShortcutListBox.Items.Clear();
            foreach(var id in ManuallyCreatedShortcuts)
            {
                var item = new ListBoxItem();
                item.Tag = id;
                item.ContextMenu = new ContextMenu();
                var menuItem = new MenuItem { Header = "Remove Entry", Tag = id};
                menuItem.Click += RemoveManual_Click;
                item.ContextMenu.Items.Add(menuItem);
                var game = plugin.PlayniteApi.Database.Games.Get(id);
                item.Content = game == null? "Game not found: " + id.ToString() : $"{game.Name} ({ShortcutSync.GetSourceName(game)})";
                item.ToolTip = item.Content;
                plugin.settingsView.ManuallyCreatedShortcutListBox.Items.Add(item);
            }
            plugin.settingsView.ExcludedGamesListBox.Items.Clear();
            foreach (var id in ExcludedGames)
            {
                var item = new ListBoxItem();
                item.Tag = id;
                item.ContextMenu = new ContextMenu();
                var menuItem = new MenuItem { Header = "Remove Entry", Tag = id };
                menuItem.Click += RemoveExcluded_Click;
                item.ContextMenu.Items.Add(menuItem);
                var game = plugin.PlayniteApi.Database.Games.Get(id);
                item.Content = game == null ? "Game not found: " + id.ToString() : $"{game.Name} ({ShortcutSync.GetSourceName(game)})";
                item.ToolTip = item.Content;
                plugin.settingsView.ExcludedGamesListBox.Items.Add(item);
            }
        }

        private void RemoveExcluded_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            List<Guid> toRemoveId = new List<Guid>();
            plugin.settingsView.ExcludedGamesListBox.Dispatcher.Invoke(() =>
            {
                var menuItem = sender as MenuItem;
                List<ListBoxItem> toRemove = new List<ListBoxItem>();
                foreach (ListBoxItem selected in plugin.settingsView.ExcludedGamesListBox.SelectedItems)
                {
                    toRemove.Add(selected);
                }
                foreach (var item in toRemove)
                {
                    item.Dispatcher.Invoke(() => toRemoveId.Add((Guid)item.Tag));
                }
                foreach (ListBoxItem item in toRemove)
                {
                    plugin.settingsView.ExcludedGamesListBox.Items.Remove(item);
                }
            });
            plugin.RemoveFromExclusionList(toRemoveId, this);
        }

        //private void AddGamesManuallyButton_Click(object sender, System.Windows.RoutedEventArgs e)
        //{
        //    List<GenericItemOption> items = (from game in plugin.PlayniteApi.Database.Games select new GenericItemOption(game.Name, game.Id.ToString())).ToList();
        //    var selected = plugin.PlayniteApi.Dialogs.ChooseItemWithSearch(
        //        items,
        //        (filter) => (from item in items where item.Name.ToLower().Contains(filter.ToLower()) select item).ToList()
        //    );
        //}

        private void RemoveManual_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            List<Guid> toRemoveId = new List<Guid>();
            plugin.settingsView.ManuallyCreatedShortcutListBox.Dispatcher.Invoke(() => 
            {
                var menuItem = sender as MenuItem;
                List<ListBoxItem> toRemove = new List<ListBoxItem>();
                foreach (ListBoxItem selected in plugin.settingsView.ManuallyCreatedShortcutListBox.SelectedItems)
                {
                    toRemove.Add(selected);
                }
                foreach (var item in toRemove)
                {
                    item.Dispatcher.Invoke(() => toRemoveId.Add((Guid)item.Tag));
                }
                foreach (ListBoxItem item in toRemove)
                {
                    plugin.settingsView.ManuallyCreatedShortcutListBox.Items.Remove(item);
                }
            });
            plugin.RemoveShortcutsManually(toRemoveId, this);
        }

        // Callback for the Select Folder Button
        private void Bt_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string path = plugin.PlayniteApi.Dialogs.SelectFolder();
            if (!string.IsNullOrEmpty(path))
            {
                plugin.settingsView.PathTextBox.Text = path;
            }
        }

        public void CancelEdit()
        {
            // Code executed when user decides to cancel any changes made since BeginEdit was called.
            // This method should revert any changes made to Option1 and Option2.
        }

        public void EndEdit()
        {
            var container = plugin.settingsView.SourceSettingsStack;
            foreach (CheckBox checkBox in container.Children)
            {
                SourceOptions[checkBox.Tag as string] = (bool)checkBox.IsChecked;
            }
            container = plugin.settingsView.PlayActionSettingsStack;
            foreach (CheckBox checkBox in container.Children)
            {
                EnabledPlayActions[checkBox.Tag as string] = (bool)checkBox.IsChecked;
            }
            // Code executed when user decides to confirm changes made since BeginEdit was called.
            plugin.SavePluginSettings(this);
            OnSettingsChanged?.Invoke();
        }

        public bool VerifySettings(out List<string> errors)
        {
            // Code execute when user decides to confirm changes made since BeginEdit was called.
            // Executed before EndEdit is called and EndEdit is not called if false is returned.
            // List of errors is presented to user if verification fails.
            var tempShortcutPath = plugin.settingsView.PathTextBox.Text;
            var bt = plugin.settingsView.SelectFolderButton;
            bt.Click -= Bt_Click;
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
            ShortcutSync.logger.Info($"Verified Settings. Old {ShortcutPath}. New {tempShortcutPath}");
            if (ShortcutPath != tempShortcutPath)
            {
                OnPathChanged?.Invoke(ShortcutPath, tempShortcutPath);
                ShortcutPath = tempShortcutPath;
            }
            return true;
        }
    }
}