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
                        UpdateShortcuts(PlayniteApi.Database.Games);
                    }),
                new ExtensionFunction(
                    "Force Update Selected Shortcuts",
                    () =>
                    {
                        UpdateShortcuts(PlayniteApi.MainView.SelectedGames);
                    })
            };
        }

        public override void OnApplicationStarted()
        {
            Directory.CreateDirectory(settings.ShortcutPath);
            if (settings.UpdateOnStartup)
            {
                UpdateShortcuts(PlayniteApi.Database.Games);
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
            UpdateShortcuts(e.AddedItems.Union(e.RemovedItems));
        }

        /// <summary>
        /// Callback handling game item updates. Updates the shortcuts 
        /// of all changed games.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Games_ItemUpdated(object sender, ItemUpdatedEventArgs<Game> e)
        {
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
                        UpdateShortcuts(from update in e.UpdatedItems where !WasLaunchedOrClosed(update) select update.NewData);
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
        public UpdateStatus UpdateShortcut(Game game, bool forceUpdate, bool multiple = false)
        {
            UpdateStatus status = UpdateStatus.None;
            string path = GetShortcutPath(game, includeSourceName: multiple);
            // determine whether to create/update the shortcut or to delete it
            bool shortcutExists = System.IO.File.Exists(path);
            bool keepShortcut = ShouldKeepShortcut(game);
            // Keep/Update shortcut
            if (keepShortcut)
            {
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
                if (RemoveShortcut(game))
                {
                    status = UpdateStatus.Deleted;
                }
            }
            return status;
        }

        private bool ShouldKeepShortcut(Game game)
        {
            settings.SourceOptions.TryGetValue(game.Source.Name, out bool sourceEnabled);
            bool exludeBecauseHidden = settings.ExcludeHidden && game.Hidden;
            bool keepShortcut =
              (game.IsInstalled ||
               !settings.InstalledOnly) &&
               sourceEnabled &&
               !exludeBecauseHidden;
            return keepShortcut;
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
        private string GetShortcutPath(Game game, bool includeSourceName = true, string extension = ".lnk")
        {
            var validName = GetSafeFileName(game.Name);
            string path;
            if (includeSourceName)
            {
                path = Path.Combine(settings.ShortcutPath, validName + " (" + game.Source.Name + ")" + extension);
            } else
            {
                path = Path.Combine(settings.ShortcutPath, validName + extension);
            }
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
        /// Calls update shortcuts on all games in a seperate thread. 
        /// Depending on the settings it removes 
        /// deleted games.
        /// </summary>
        /// <param name="gamesToUpdate">Games to update.</param>
        public void UpdateShortcuts(IEnumerable<Game> gamesToUpdate)
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

                Dictionary<string, Tuple<IList<Game>, IList<Game>>> pathDict = new Dictionary<string, Tuple<IList<Game>, IList<Game>>>();
                foreach (var game in gamesToUpdate)
                {
                    string path = GetShortcutPath(game: game, includeSourceName: false);

                    if (!pathDict.ContainsKey(path))
                    {
                        pathDict.Add(path, 
                            new Tuple<IList<Game>, IList<Game>>(new List<Game>(), new List<Game>()));
                    } 
                    if (ShouldKeepShortcut(game))
                    {
                        pathDict[path].Item1.Add(game);
                    } else
                    {
                        pathDict[path].Item2.Add(game);
                    }
                }
                foreach (var games in pathDict.Values)
                {
                    bool useSourceName = games.Item1.Count > 1;
                    string gameName;
                    if (games.Item1.Count > 0)
                    {
                        gameName = games.Item1[0].Name;
                    } else
                    {
                        gameName = games.Item2[0].Name;
                    }
                    var existingShortcuts = GetExistingShortcuts(settings.ShortcutPath, GetSafeFileName(gameName));
                    int copyCount = existingShortcuts.Count;
                    useSourceName = copyCount > 1 || useSourceName;
                    foreach (var game in games.Item1.Union(games.Item2))
                    {
                        if (existingShortcuts.TryGetValue(game.Id, out string file))
                        {
                            bool includesSourceName = file == GetShortcutPath(game, true);
                            if (!ShouldKeepShortcut(game))
                            {
                                System.IO.File.Delete(file);
                            } else 
                            if (!useSourceName && includesSourceName)
                            {
                                string newPath = GetShortcutPath(game, false);
                                if (System.IO.File.Exists(newPath)) System.IO.File.Delete(newPath);
                                System.IO.File.Move(file, newPath);
                            } else if (useSourceName && !includesSourceName)
                            {
                                string newPath = GetShortcutPath(game, true);
                                if (System.IO.File.Exists(newPath)) System.IO.File.Delete(newPath);
                                System.IO.File.Move(file, newPath);
                            }
                        } 
                        string shortcutPath = GetShortcutPath(game: game, includeSourceName: useSourceName);
                        switch (UpdateShortcut(game, settings.ForceUpdate, useSourceName))
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
                }
                // LogInfo($"Updated: {updated}, created: {created}, removed: {removed}");
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
                description: "Launch " + game.Name + " on " + game.Source.Name + "." + $" [{game.Id}]",
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
                description: "Launch " + game.Name + " on " + game.Source.Name + "." + $" [{game.Id}]",
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
                description: "Launch " + game.Name + " on " + game.Source.Name + " via Playnite." + $" [{game.Id}]");
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
                description: "Launch " + game.Name + " on " + game.Source.Name + "." + $" [{game.Id}]");
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

        private Dictionary<Guid, string> 
            GetExistingShortcuts(string folderPath, string shortcutName = "")
        {
            var games = new Dictionary<Guid, string>();
            var files = Directory.GetFiles(folderPath, shortcutName + "*.lnk");
            foreach(var file in files)
            {
                var shortcut = OpenLnk(file);
                string desc = shortcut.Description;
                int startId = -1, endId = -1;
                for (int i = desc.Length-1; i >= 0; --i )
                {
                    if (desc[i] == ']')
                    {
                        endId = i - 1;
                    } else if (desc[i] == '[')
                    {
                        startId = i + 1;
                        break;
                    }
                }
                if (startId != -1 && endId != -1)
                {
                    Guid id = Guid.Parse(desc.Substring(startId, endId - startId + 1));
                    Game game = PlayniteApi.Database.Games[id];
                    if (games.ContainsKey(game.Id))
                    {
                        System.IO.File.Delete(file);
                    } else
                    {
                        games.Add(game.Id, file);
                    }
                }
            }
            return games;
        }

        public IWshShortcut OpenLnk(string shortcutPath)
        {
            if (System.IO.File.Exists(shortcutPath))
            {
                try
                {
                    var shell = new WshShell();
                    var shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                    return shortcut;
                }
                catch (Exception ex)
                {
                    LogError($"Could not open shortcut at \"{shortcutPath}\". Exception: {ex.Message}");
                }
            }
            return null;
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