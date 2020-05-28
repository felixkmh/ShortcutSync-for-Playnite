using IWshRuntimeLibrary;
using Playnite.SDK;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ShortcutSync
{
    public class ShortcutSync : Plugin
    {
        private static readonly ILogger logger = LogManager.GetLogger();
        private Thread thread;
        private ShortcutSyncSettings settings { get; set; }
        public ShortcutSyncSettingsView settingsView { get; set; }

        public override Guid Id { get; } = Guid.Parse("8e48a544-3c67-41f8-9aa0-465627380ec8");

        public enum UpdateStatus
        {
            None,
            Updated,
            Deleted,
            Created
        }

        public ShortcutSync(IPlayniteAPI api) : base(api)
        {
            settings = new ShortcutSyncSettings(this);
        }

        public override IEnumerable<ExtensionFunction> GetFunctions()
        {
            return new List<ExtensionFunction>
            {
                new ExtensionFunction(
                    "Update All Shortcuts",
                    () =>
                    {
                        // Update shortcuts of all (installed) games
                        UpdateShortcuts();
                    }),
                new ExtensionFunction(
                    "Force Update Selected Shortcuts",
                    () =>
                    {
                        thread?.Join();
                        thread = new Thread(() =>
                        {
                            PlayniteApi.Database.Games.BeginBufferUpdate();
                            int removed = 0;
                            int updated = 0;
                            int created = 0;
                            foreach (var game in PlayniteApi.MainView.SelectedGames)
                            {
                                // Update shortcuts of selected games
                                switch (UpdateShortcut(game, true))
                                {
                                    case UpdateStatus.Updated:
                                        updated++;
                                        break;
                                    case UpdateStatus.Deleted:
                                        removed++;
                                        break;
                                    case UpdateStatus.Created:
                                        created++;
                                        break;
                                    default:
                                        break;
                                }
                            }
                            LogInfo($"Updated: {updated}, created: {created}, removed: {removed}");
                            PlayniteApi.Database.Games.EndBufferUpdate();
                        });
                        thread.Start();
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
                UpdateShortcut(game, settings.ForceUpdate);
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
            foreach (var change in e.UpdatedItems)
            {
                // Don't update shortcut if the game was only closed or launched
                if (WasLaunchedOrClosed(change))
                {
                    return;
                }

                ThreadPool.QueueUserWorkItem((_) =>
                {
                    bool success = false;
                    for (int i = 0; i < 10; ++i)
                    {
                        // Workaround because icon files are 
                        // still locked when ItemUpdated is called
                        Thread.Sleep(10);
                        try
                        {
                            UpdateShortcut(change.NewData, settings.ForceUpdate);
                            success = true;
                        }
                        catch (Exception)
                        {
                            LogManager.GetLogger().Debug($"Could not convert icon. Trying again...");
                        }
                        if (success) break;
                    }
                });

            }
        }

        public override void OnApplicationStopped()
        {
            // Unsubscribe from library change events.
            PlayniteApi.Database.Games.ItemUpdated -= Games_ItemUpdated;
            PlayniteApi.Database.Games.ItemCollectionChanged -= Games_ItemCollectionChanged;
            thread?.Join();
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
        /// Update, create or remove a shortcut to a game.
        /// </summary>
        /// <param name="game">Game the shortcut should point to.</param>
        /// <param name="forceUpdate">If true, existing shortcut will be overwritten.</param>
        /// <returns>Status of the associated shortcut.</returns>
        public UpdateStatus UpdateShortcut(Game game, bool forceUpdate)
        {
            UpdateStatus status = UpdateStatus.None;
            string path = GetShortcutPath(game);
            settings.SourceOptions.TryGetValue(game.Source.Name, out bool sourceEnabled);
            // determine whether to create/update the shortcut or to delete it
            bool exludeBecauseHidden = settings.ExcludeHidden && game.Hidden;
            bool keepShortcut = 
               (game.IsInstalled  || 
                !settings.InstalledOnly) &&
                sourceEnabled &&
                !exludeBecauseHidden;
            // Keep/Update shortcut
            if (keepShortcut)
            {
                bool shortcutExists = System.IO.File.Exists(path);
                bool willCreateShortcut = !shortcutExists || forceUpdate;
                if (willCreateShortcut)
                {
                    // Assign status accordingly
                    if (shortcutExists)
                    {
                        status = UpdateStatus.Updated;
                    } else
                    {
                        status = UpdateStatus.Created;
                    }
                    CreateShortcut(game, path);
                }

            }
            // Remove shortcut
            else
            {
                RemoveShortcut(game);
                status = UpdateStatus.Deleted;
            }
            return status;
        }

        /// <summary>
        /// Create shortcut for a game at a given location.
        /// </summary>
        /// <param name="game">Game the shortcut should launch.</param>
        /// <param name="shortcutPath">Full path to the shortcut location.</param>
        public void CreateShortcut(Game game, string shortcutPath)
        {
            string icon = string.Empty;

            if (!string.IsNullOrEmpty(game.Icon))
            {
                icon = PlayniteApi.Database.GetFullFilePath(game.Icon);
            }

            if (System.IO.File.Exists(icon))
            {
                if (Path.GetExtension(icon) != ".ico")
                {
                    if (ConvertToIcon(icon, out string output))
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
            // Creat playnite URI if game is not installed
            // or if Use PlayAction option is disabled
            if (!game.IsInstalled || !settings.UsePlayAction)
            {
                CreateLnkURL(shortcutPath, icon, game);
            } 
            else 
            {
                if (game.Source.Name == "Xbox")
                {
                    CreateLnkFileXbox(shortcutPath, icon, game);
                }
                else
                {
                    if (game.PlayAction.Type == GameActionType.URL)
                    {
                        CreateLnkURLDirect(shortcutPath, icon, game);
                    }
                    else
                    {
                        CreateLnkFile(shortcutPath, icon, game);
                    }
                }
            }
        }

        /// <summary>
        /// Converts an image into an icon in the same folder as the 
        /// source image naming it with its MD5 hash.
        /// Mostly taken from the ConvertToIcon implementation
        /// in Playnite.Common
        /// </summary>
        /// <param name="icon">Full path to the source image.</param>
        /// <param name="output">Full path to the output file on success.</param>
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
                LogError("Could not create temporary file. Exception: " + ex.Message);
                return false;
            }
            bool success = false;
            try
            {
                success = BitmapExtensions.ConvertToIcon(icon, tempIcon);
            }
            catch (Exception ex)
            {
                LogManager.GetLogger().Debug($"Could not convert icon {icon}. Exception: " + ex.Message);
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
                    if (System.IO.File.Exists(newPath))
                    {
                        System.IO.File.Delete(tempIcon);
                    }
                    else
                    {
                        System.IO.File.Move(tempIcon, newPath);
                    }
                }
                catch (Exception ex)
                {
                    LogError($"Could not move converted icon to \"{newPath}\". Exception: {ex.Message}");
                }
                output = newPath;
            }
            else
            {
                LogError($"Could not convert file \"{icon}\" to ico.");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Creates the shortcut path based on the 
        /// game's name and the path in the settings.
        /// </summary>
        /// <param name="game">The game the shortcut will point to.</param>
        /// <param name="extension">Filename extension including ".".</param>
        /// <returns>
        /// String containing path and file name for the .url shortcut.
        /// </returns>
        private string GetShortcutPath(Game game, string extension = ".lnk")
        {
            var validName = GetSafeFileName(game.Name);
            var path = Path.Combine(settings.ShortcutPath, validName + " (" + game.Source.Name + ")" + extension);
            return path;
        }

        /// <summary>
        /// Tries to remove the shortcut for a given game
        /// if it exists.
        /// </summary>
        /// <param name="game">The game the shortcut points to.</param>
        private bool RemoveShortcut(Game game)
        {
            var path = GetShortcutPath(game);
            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
                return true;
            } else
            {
                return false;
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
            // Updatind all games can take some time
            // execute in a task so main thread is not blocked.
            thread?.Join();
            thread = new Thread(() =>
            {
                // Buffer updates while updating shortcuts
                // changes during this process will be handled
                // by the events afterwards
                PlayniteApi.Database.Games.BeginBufferUpdate();
                int removed = 0;
                int updated = 0;
                int created = 0;
                foreach (var game in PlayniteApi.Database.Games)
                {
                    switch (UpdateShortcut(game, settings.ForceUpdate))
                    {
                        case UpdateStatus.Updated:
                            updated++;
                            break;
                        case UpdateStatus.Deleted:
                            removed++;
                            break;
                        case UpdateStatus.Created:
                            created++;
                            break;
                        default:
                            break;
                    }
                }
                LogInfo($"Updated: {updated}, created: {created}, removed: {removed}");
                PlayniteApi.Database.Games.EndBufferUpdate();
            });
            thread.Start();
        }

        /// <summary>
        /// Checks whether a game entry changes because it
        /// was launched or closed.
        /// </summary>
        /// <param name="data">The event data of the changed game.</param>
        /// <returns></returns>
        private static bool WasLaunchedOrClosed(ItemUpdateEvent<Game> data)
        {
            return ( data.NewData.IsRunning && !data.OldData.IsRunning) || ( data.NewData.IsLaunching && !data.OldData.IsLaunching) ||
                   (!data.NewData.IsRunning &&  data.OldData.IsRunning) || (!data.NewData.IsLaunching &&  data.OldData.IsLaunching);
        }

        /// <summary>
        /// Creates a .lnk shortcut given a game with a File PlayAction.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="iconPath"></param>
        /// <param name="game"></param>
        public void CreateLnkFile(string path, string iconPath, Game game)
        {
            string workingDirectory = PlayniteApi.ExpandGameVariables(game, game.PlayAction.WorkingDir);
            string targetPath = PlayniteApi.ExpandGameVariables(game, game.PlayAction.Path);
            if (!string.IsNullOrEmpty(workingDirectory))
            {
                targetPath = Path.Combine(workingDirectory, targetPath);
            }
            CreateLnk(
                shortcutPath: path,
                targetPath: targetPath,
                iconPath: iconPath,
                description: "Launch " + game.Name + " on " + game.Source.Name + ".",
                workingDirectory: workingDirectory,
                arguments: game.PlayAction.Arguments);
        }

        /// <summary>
        /// Creates a .lnk shortcut given a game with a File PlayAction.
        /// Specialized for Windows UWP apps.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="iconPath"></param>
        /// <param name="game"></param>
        public void CreateLnkFileXbox(string path, string iconPath, Game game)
        {
            string targetPath = PlayniteApi.ExpandGameVariables(game, game.PlayAction.Path);
            CreateLnk(
                shortcutPath: path,
                targetPath: @"C:Windows\explorer.exe",
                iconPath: iconPath,
                description: "Launch " + game.Name + " on " + game.Source.Name + ".",
                workingDirectory: "Applications",
                arguments: game.PlayAction.Arguments);
        }
        /// <summary>
        /// Creates a .lnk shortcut launching a game using a playnite:// url.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="iconPath"></param>
        /// <param name="game"></param>
        public void CreateLnkURL(string path, string iconPath, Game game)
        {
            CreateLnk(
                shortcutPath: path,
                targetPath: $"playnite://playnite/start/{game.Id}",
                iconPath: iconPath,
                description: "Launch " + game.Name + " on " + game.Source.Name + " via Playnite.");
        }
        /// <summary>
        /// Creates a .lnk shortcut given a game with a URL PlayAction.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="iconPath"></param>
        /// <param name="game"></param>
        public void CreateLnkURLDirect(string path, string iconPath, Game game)
        {
            CreateLnk(
                shortcutPath: path,
                targetPath: game.PlayAction.Path,
                iconPath: iconPath,
                description: "Launch " + game.Name + " on " + game.Source.Name + ".");
        }

        /// <summary>
        /// Create a lnk format shortcut.
        /// </summary>
        /// <param name="shortcutPath">Full path and filename of the shortcut file.</param>
        /// <param name="targetPath">Full target path of the shortcut.</param>
        /// <param name="iconPath">Full path to the icon for the shortcut.</param>
        /// <param name="description">Description of the shortcut.</param>
        /// <param name="workingDirectory">Optional path to the working directory.</param>
        /// <param name="arguments">Optional launch argurments.</param>
        public void CreateLnk(
            string shortcutPath, string targetPath, string iconPath, 
            string description, string workingDirectory = "", string arguments = "")
        {
            try
            {
                var shell = new WshShell();
                var shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.IconLocation = iconPath;
                shortcut.TargetPath = targetPath;
                shortcut.Description = description;
                shortcut.WorkingDirectory = workingDirectory;
                shortcut.Arguments = arguments;
                shortcut.Save();
            }
            catch (Exception ex)
            {
                LogError($"Could not create shortcut at \"{shortcutPath}\". Exception: {ex.Message}");
            }
        }

        #region Logging
        private enum LogType
        {
            Info,
            Error,
            Debug,
            Warn
        }

        /// <summary>
        /// Wrapper to write different messages to the log.
        /// </summary>
        /// <param name="logType">Type of the message.</param>
        /// <param name="message">The message to log.</param>
        private void Log(LogType logType, string message)
        {
            var logger = LogManager.GetLogger();
            string pre = "[Plugin: ShortcutSync] ";
            switch (logType)
            {
                case LogType.Info:
                    logger.Info(pre + message);
                    break;
                case LogType.Error:
                    logger.Error(pre + message);
                    break;
                case LogType.Debug:
                    logger.Debug(pre + message);
                    break;
                case LogType.Warn:
                    logger.Warn(pre + message);
                    break;
                default:
                    break;
            }
        }
        /// <summary>
        /// Add info log entry.
        /// </summary>
        /// <param name="message">Log message.</param>
        private void LogInfo(string message)
        {
            Log(LogType.Info, message);
        }
        /// <summary>
        /// Add error log entry.
        /// </summary>
        /// <param name="message">Log message.</param>
        private void LogError(string message)
        {
            Log(LogType.Error, message);
        }
        /// <summary>
        /// Add debug log entry.
        /// </summary>
        /// <param name="message">Log message.</param>
        private void LogDebug(string message)
        {
            Log(LogType.Debug, message);
        }
        /// <summary>
        /// Add warning log entry.
        /// </summary>
        /// <param name="message">Log message.</param>
        private void LogWarn(string message)
        {
            Log(LogType.Warn, message);
        }
        #endregion
    }
}