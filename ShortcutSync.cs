using Playnite.SDK;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Drawing.Imaging;
using System.Threading;

namespace ShortcutSync
{
    public class ShortcutSync : Plugin
    {
        private static readonly ILogger logger = LogManager.GetLogger();

        private ShortcutSyncSettings settings { get; set; }
        public ShortcutSyncSettingsView settingsView { get; set; }

        public override Guid Id { get; } = Guid.Parse("8e48a544-3c67-41f8-9aa0-465627380ec8");

        public ShortcutSync(IPlayniteAPI api) : base(api)
        {
            settings = new ShortcutSyncSettings(this);
        }

        public override IEnumerable<ExtensionFunction> GetFunctions()
        {
            return new List<ExtensionFunction>
            {
                new ExtensionFunction(
                    "Update Start Menu Shortcuts",
                    () =>
                    {
                        // Update shortcuts of all (installed) games
                        PlayniteApi.Database.Games.ItemUpdated -= Games_ItemUpdated;
                        PlayniteApi.Database.Games.ItemCollectionChanged -= Games_ItemCollectionChanged;
                        UpdateShortcuts();
                        PlayniteApi.Database.Games.ItemUpdated += Games_ItemUpdated;
                        PlayniteApi.Database.Games.ItemCollectionChanged += Games_ItemCollectionChanged;
                    })
            };
        }

        public override void OnApplicationStarted()
        {
            Directory.CreateDirectory(settings.ShortcutPath);
            if (settings.UpdateOnStartup)
            {
                UpdateShortcuts();
            }
            PlayniteApi.Database.Games.ItemUpdated += Games_ItemUpdated;
            PlayniteApi.Database.Games.ItemCollectionChanged += Games_ItemCollectionChanged;
        }

        /// <summary>
        /// Callback handling changes in the games database.
        /// Shortcuts are removed or created according to settings.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Games_ItemCollectionChanged(object sender, ItemCollectionChangedEventArgs<Game> e)
        {
            // Update or create shortcuts for newly added games
            foreach (var game in e.AddedItems)
            {
                UpdateShortcut(game);
            }
            // Delete shortcuts for games removed from the library
            // if only installed games should have shortcuts
            if (settings.InstalledOnly)
            {
                foreach (var game in e.RemovedItems)
                {
                    RemoveShortcut(game);
                }
            }
        }

        /// <summary>
        /// Callback handling game item updates. Updates the shortcuts 
        /// of all changed games.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Games_ItemUpdated(object sender, ItemUpdatedEventArgs<Game> e)
        {
            foreach (var game in e.UpdatedItems)
            {
                // keep shortcut if game is installed or settings indicate that
                // all shortcuts should be kept.
                bool keepShortcut = game.NewData.IsInstalled || !settings.InstalledOnly;
                if (keepShortcut)
                {
                    ThreadPool.QueueUserWorkItem( (_) =>
                    {
                        bool success = false;
                        for (int i = 0; i < 10; ++i)
                        {
                            // Workaround because icon files are 
                            // still locked when ItemUpdated is called
                            Thread.Sleep(10);
                            try
                            {
                                UpdateShortcut(game.NewData);
                                success = true;
                            }
                            catch (Exception ex)
                            {
                                LogManager.GetLogger().Debug($"Could not convert icon. Trying again...");
                            }
                            if (success) break;
                        }
                    });
                } else
                {
                    RemoveShortcut(game.NewData);
                }
            }
        }

        public override void OnApplicationStopped()
        {
            // Unsubscribe from library change events.
            PlayniteApi.Database.Games.ItemUpdated -= Games_ItemUpdated;
            PlayniteApi.Database.Games.ItemCollectionChanged -= Games_ItemCollectionChanged;
        }


        public override ISettings GetSettings(bool firstRunSettings)
        {
            return settings;
        }

        public override UserControl GetSettingsView(bool firstRunSettings)
        {
            settingsView = new ShortcutSyncSettingsView();
            return settingsView;
        }

        /// <summary>
        /// Update or create a shortcut to a game.
        /// </summary>
        /// <param name="game">Game the shortcut should point to.</param>
        public void UpdateShortcut(Game game)
        {
            settings.SourceOptions.TryGetValue(game.Source.Name, out bool createShortcut);
            // Early exit if shortcuts disabled for this game's source
            // or source is not set in options.
            if (!createShortcut)
            {
                RemoveShortcut(game);
                return;
            }

            string path = GetShortcutPath(game);
            // determine whether to create/update the shortcut or to delete it
            bool keepShortcut = game.IsInstalled || !settings.InstalledOnly;
            // Keep/Update shortcut
            if (keepShortcut)
            {
                // If the shortcut already exists,
                // exit if forceUpdate is disabled
                if (File.Exists(path) && !settings.ForceUpdate)
                {
                    return;
                }

                string icon = string.Empty;

                if (!string.IsNullOrEmpty(game.Icon))
                {
                    icon = PlayniteApi.Database.GetFullFilePath(game.Icon);
                }

                if (File.Exists(icon))
                {
                    if (Path.GetExtension(icon) != ".ico")
                    {
                        if(ConvertToIcon(icon, out string output))
                        {
                            icon = output;
                        }
                    }
                }
                // If no icon was found, use a fallback icon.
                else if (String.IsNullOrEmpty(icon))
                {
                    icon = Path.Combine(PlayniteApi.Paths.ApplicationPath, "Playnite.DesktopApp.exe");
                }

                // Build shortcut string
                StringBuilder url = new StringBuilder();
                url.AppendLine("[InternetShortcut]");
                url.AppendLine("IconIndex=0");
                url.Append("IconFile=").AppendLine(icon);
                url.Append("URL=").AppendLine($"playnite://playnite/start/{game.Id}");
                File.WriteAllText(path, url.ToString());
            }
            // Remove shortcut
            else
            {
                if (File.Exists(path)) RemoveShortcut(game);
            }

        }

        /// <summary>
        /// Converts an image into an icon in the same folder as the 
        /// source image naming it with its MD5 hash.
        /// Mostly taken from the ConvertToIcon implementation
        /// in Playnite.Common
        /// </summary>
        /// <param name="icon">Full path to the source image to convert.</param>
        /// <returns>Whether conversion succeded.</returns>
        private bool ConvertToIcon(string icon, out string output)
        {
            output = string.Empty;
            string tempIcon;
            // Try to create temp file to write the icon to
            try
            {
                tempIcon = Path.GetTempFileName();
            }
            catch (IOException ex)
            {
                PlayniteApi.Dialogs.ShowErrorMessage(ex.Message, "Could not create temporary file.");
                return false;
            }
            bool success = false;
            try
            {
                success = BitmapExtensions.ConvertToIcon(icon, tempIcon);
            }
            catch (Exception ex)
            {
                LogManager.GetLogger().Debug($"Could not convert icon {icon}");
                throw ex;
            }
            if (success)
            {
                var md5 = Playnite.Common.FileSystem.GetMD5(tempIcon);
                var newPath = Path.Combine(Path.GetDirectoryName(icon), md5 + ".ico");
                // Move new icon into image folder, possibly overwriting previous icon
                // if it was the same icon or was converted before
                try
                {
                    if (File.Exists(newPath))
                    {
                        File.Delete(tempIcon);
                    } else
                    {
                        File.Move(tempIcon, newPath);
                    }
                }
                catch (Exception ex)
                {
                    PlayniteApi.Dialogs.ShowErrorMessage(ex.Message, "Could not move converted icon");
                }
                output = newPath;
            } else
            {
                PlayniteApi.Dialogs.ShowErrorMessage($"Could not file {icon} convert to ico.", "ShortcutSync");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Creates the shortcut path based on the 
        /// game's name and the path in the settings.
        /// </summary>
        /// <param name="game">The game the shortcut will point to.</param>
        /// <returns>
        /// String containing path and file name for the .url shortcut.
        /// </returns>
        private string GetShortcutPath(Game game)
        {
            var validName = GetSafeFileName(game.Name);
            var path = Path.Combine(settings.ShortcutPath, validName + ".url");
            return path;
        }

        /// <summary>
        /// Tries to remove the shortcut for a given game
        /// if it exists.
        /// </summary>
        /// <param name="game">The game the shortcut points to.</param>
        private void RemoveShortcut(Game game)
        {
            var path = GetShortcutPath(game);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        /// <summary>
        /// Removes characters that are not valid in 
        /// paths and filenames.
        /// </summary>
        /// <param name="validName">string containing a possibly invalid path.</param>
        /// <returns>String containing the valid path and filename.</returns>
        private static string GetSafeFileName(string validName)
        {
            foreach (var c in Path.GetInvalidFileNameChars()) validName = validName.Replace(c.ToString(), "");
            return validName;
        }

        /// <summary>
        /// Calls update shortcuts on all games. 
        /// Depending on the settings it removes 
        /// deleted games.
        /// </summary>
        public void UpdateShortcuts()
        {
            foreach (var game in PlayniteApi.Database.Games.Where(g => g.IsInstalled || !settings.InstalledOnly))
            {
                UpdateShortcut(game);
            }
            if (settings.InstalledOnly)
            {
                foreach (var game in PlayniteApi.Database.Games.Where(g => !g.IsInstalled))
                {
                    RemoveShortcut(game);
                }
            }
        }
    }
}