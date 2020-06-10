using IWshRuntimeLibrary;
using Playnite.Common;
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
        private Dictionary<Guid, string> existingShortcuts { get; set; }
        private Dictionary<string, IList<Guid>> shortcutNameToGameId { get; set; } 

        public override Guid Id { get; } = Guid.Parse("8e48a544-3c67-41f8-9aa0-465627380ec8");

        public enum UpdateStatus
        {
            None,
            Updated,
            Deleted,
            Created,
            Error
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
                        (existingShortcuts, shortcutNameToGameId) = GetExistingShortcuts(settings.ShortcutPath);
                        // Update shortcuts of all (installed) games
                        UpdateShortcuts(PlayniteApi.Database.Games);
                    }),
                new ExtensionFunction(
                    "Force Update Selected Shortcuts",
                    () =>
                    {
                        (existingShortcuts, shortcutNameToGameId) = GetExistingShortcuts(settings.ShortcutPath);
                        UpdateShortcuts(PlayniteApi.MainView.SelectedGames, true);
                    })
            };
        }

        public override void OnApplicationStarted()
        {
            try
            {
                Directory.CreateDirectory(settings.ShortcutPath);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Could not create directory \"{settings.ShortcutPath}\". Try choosing another folder.");
            }
            (existingShortcuts, shortcutNameToGameId) = GetExistingShortcuts(settings.ShortcutPath);
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
                        logger.Debug($"Could not convert icon. Trying again...");
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
        /// <param name="multiple">Specifies whether source name should be included in the shortcut
        /// name to prevent conflicts if multiple games have the same shortcut name.</param>
        /// <returns>Status of the associated shortcut.</returns>
        public UpdateStatus UpdateShortcut(Game game, bool forceUpdate, bool multiple = false)
        {
            UpdateStatus status = UpdateStatus.None;
            // determine whether to create/update the shortcut or to delete it
            bool shortcutExists = existingShortcuts.TryGetValue(game.Id, out string currentPath);
            string desiredPath = GetShortcutPath(game, multiple);
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
                        if (desiredPath != currentPath)
                        {
                            try
                            {
                                System.IO.File.Move(currentPath, desiredPath);
                            }
                            catch (Exception ex)
                            {
                                logger.Error(ex, $"Could not move \"{currentPath}\" to \"{desiredPath}\".");
                                return UpdateStatus.Error;
                            }
                        }
                    } else
                    {
                        status = UpdateStatus.Created;
                    }
                    CreateShortcut(game, desiredPath);
                    existingShortcuts[game.Id] = desiredPath;
                    string safeGameName = GetSafeFileName(game.Name).ToLower();
                    if (!shortcutNameToGameId.ContainsKey(safeGameName))
                    {
                        shortcutNameToGameId[safeGameName] = new List<Guid>();
                    }
                    shortcutNameToGameId[safeGameName].AddMissing(game.Id);
                }

            }
            // Remove shortcut
            else
            {
                if (RemoveShortcut(game))
                {
                    status = UpdateStatus.Deleted;
                    existingShortcuts.Remove(game.Id);
                    if (shortcutNameToGameId.TryGetValue(GetSafeFileName(game.Name).ToLower(), out var games))
                    {
                        games.Remove(game.Id);
                    }
                }
            }
            return status;
        }


        /// <summary>
        /// Checks whether a game's shortcut will be kept when updated
        /// based on its state and the plugin settings.
        /// </summary>
        /// <param name="game">The game that is ckecked.</param>
        /// <returns>Whether shortcut should be kept.</returns>
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
                    // Check if icon has already been converted before
                    var iconFiles = new string[0];
                    try
                    {
                        iconFiles = System.IO.Directory.GetFiles(PlayniteApi.Database.GetFileStoragePath(game.Id), "*.ico");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Could not open folder {PlayniteApi.Database.GetFileStoragePath(game.Id)} to look for existing .ico files.");
                    }
                    if (iconFiles.Length > 0)
                    {
                        icon = iconFiles[0];
                    } 
                    else if (ConvertToIcon(icon, out string output))
                    {
                        icon = output;
                    }
                }
            }
            // If no icon was found, use a fallback icon.
            else 
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
                logger.Error(ex, "Could not create temporary file.");
                // LogError("Could not create temporary file. Exception: " + ex.Message);
                return false;
            }
            bool success = false;
            try
            {
                success = BitmapExtensions.ConvertToIcon(icon, tempIcon);
            }
            catch (Exception ex)
            {
                logger.Debug(ex, $"Could not convert icon {icon}.");
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
                    logger.Error(ex, $"Could not move converted icon to \"{newPath}\".");
                    // LogError($"Could not move converted icon to \"{newPath}\". Exception: {ex.Message}");
                }
                output = newPath;
            }
            else
            {
                logger.Error($"Could not convert file \"{icon}\" to ico.");
                // LogError($"Could not convert file \"{icon}\" to ico.");
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
            if (existingShortcuts.TryGetValue(game.Id, out string path))
            {
                if (System.IO.File.Exists(path))
                {
                    try
                    {
                        System.IO.File.Delete(path);
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Could not delete shortcut at \"{path}\".");
                        return false;
                    }
                    return true;
                } else
                {
                    return false;
                }
            }
            return false;
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
        public void UpdateShortcuts(IEnumerable<Game> gamesToUpdate, bool forceUpdate = false)
        {
            bool errorOccured = false;
            // Updatind all games can take some time
            // execute in a task so main thread is not blocked.
            thread?.Join();
            thread = new Thread(() =>
            {
                // Buffer updates while updating shortcuts
                // changes during this process will be handled
                // by the events afterwards
                PlayniteApi.Database.Games.BeginBufferUpdate();

                foreach (var game in gamesToUpdate)
                {
                    string path = GetShortcutPath(game: game, includeSourceName: false);
                    string safeGameName = GetSafeFileName(game.Name);

                    List<Game> existing = new List<Game>();
                    if (shortcutNameToGameId.ContainsKey(safeGameName.ToLower()))
                    {
                        existing.AddRange(
                            from gameId 
                            in shortcutNameToGameId[safeGameName.ToLower()] 
                            where gameId != game.Id 
                            select PlayniteApi.Database.Games.Get(gameId));
                    }
                    if (ShouldKeepShortcut(game))
                    {
                        foreach (var copy in existing)
                        {
                            string newPath = GetShortcutPath(game: copy, includeSourceName: true);
                            if (existingShortcuts.TryGetValue(copy.Id, out string currentPath))
                            {
                                if (System.IO.Path.GetFileName(currentPath).ToLower() != System.IO.Path.GetFileName(newPath).ToLower())
                                {
                                    try
                                    {
                                        System.IO.File.Move(currentPath, newPath);
                                        existingShortcuts[copy.Id] = newPath;
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.Error(ex, $"Could not move {currentPath} to {newPath}, while updating {game.Name} on {game.Source.Name}.");
                                    }
                                }
                            }
                        }
                        if (UpdateShortcut(game, forceUpdate, existing.Count > 0) == UpdateStatus.Error)
                        {
                            errorOccured = true;
                        }
                    } else if (existing.Count == 1)
                    {
                        if (UpdateShortcut(game, forceUpdate, false) == UpdateStatus.Error)
                        {
                            errorOccured = true;
                        }
                        string newPath = GetShortcutPath(game: existing[0], includeSourceName: false);
                        string currentPath = existingShortcuts[existing[0].Id];
                        if (System.IO.Path.GetFileName(currentPath).ToLower() != System.IO.Path.GetFileName(newPath).ToLower())
                        {
                            if (!System.IO.File.Exists(newPath))
                            {
                                try
                                {
                                    System.IO.File.Move(currentPath, newPath);
                                }
                                catch (Exception ex)
                                {
                                    logger.Error(ex, $"Could not move {currentPath} to {newPath}, while updating {game.Name} on {game.Source.Name}.");
                                }
                                existingShortcuts[existing[0].Id] = newPath;
                            }
                        }
                    } else
                    {
                        if (UpdateShortcut(game, forceUpdate, false) == UpdateStatus.Error)
                        {
                            errorOccured = true;
                        }
                    }
                      
                }
                if (errorOccured)
                {
                    // Refresh dictionaries to properly reflect current state.
                    (existingShortcuts, shortcutNameToGameId) = GetExistingShortcuts(settings.ShortcutPath);
                }
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
                logger.Error(ex, $"Could not create shortcut at \"{shortcutPath}\".");
            }
        }

        /// <summary>
        /// Looks for a <see cref="Guid"/> inside []-brackets inside a string.
        /// </summary>
        /// <param name="description">string containing the <see cref="Guid"/>.</param>
        /// <param name="gameId">The extracted <see cref="Guid"/>.</param>
        /// <returns>Whether a valid <see cref="Guid"/> could be parsed.</returns>
        private bool ExtractIdFromLnkDescription(string description, out Guid gameId)
        {
            int startId = -1, endId = -1;
            for (int i = description.Length - 1; i >= 0; --i)
            {
                if (description[i] == ']')
                {
                    endId = i - 1;
                }
                else if (description[i] == '[')
                {
                    startId = i + 1;
                    break;
                }
            }
            if (startId != -1 && endId != -1)
            {
                return Guid.TryParse(description.Substring(startId, endId - startId + 1), out gameId);
            } else
            {
                gameId = default;
                return false;
            }
        }

        /// <summary>
        /// Searches a folder for existing shortcuts by reading
        /// their description and looking for a <see cref="Guid"/> inside []-brackets
        /// and creates dictionaries holding the results.
        /// </summary>
        /// <param name="folderPath">Full path to the folder to check for shortcuts.</param>
        /// <param name="shortcutName">Pattern to look for that is prepended to "*.lnk" to look for files.</param>
        /// <returns>A dictionary that assigns each found gameId to the full shortcut path
        /// and a dictionary assigning file names in lower case to a list of gameIds 
        /// that would create a shortcut with that name.</returns>
        private (Dictionary<Guid, string> guidToShortcut, Dictionary<string, IList<Guid>> nameToGuid)
            GetExistingShortcuts(string folderPath, string shortcutName = "")
        {
            var games = new Dictionary<Guid, string>();
            var nameToId = new Dictionary<string, IList<Guid>>();
            string[] files = new string[0];
            try
            {
                files = Directory.GetFiles(folderPath, shortcutName + "*.lnk");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Could not open folder at {folderPath}. Can not update which shortcuts already exist.");
            }
            foreach(var file in files)
            {
                var shortcut = OpenLnk(file);
                if (shortcut == null)
                {
                    // skip if shortcut could not be opened
                    logger.Warn($"Could not open shortcut at \"{file}\". Could be corrupted or file access might be restricted.");
                    continue;
                }
                if (ExtractIdFromLnkDescription(shortcut.Description, out Guid gameId))
                {
                    Game game = PlayniteApi.Database.Games.Get(gameId);
                    if (game == null)
                    {
                        logger.Warn($"Shortcut at \"{file}\" conatins Guid \"{gameId}\", but this Id is not contained in the game database.");
                        // skip if Guid was not found in game databas
                        continue;
                    }
                    if (games.ContainsKey(game.Id))
                    {
                        logger.Warn($"Shortcut at \"{file}\" is a duplicate for {game.Name} with gameId {game.Id}. Will be deleted.");
                        try
                        {
                            System.IO.File.Delete(file);
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex, $"Could not remove shortcut at \"{file}\".");
                        }
                    } else
                    {
                        games.Add(game.Id, file);
                        string safeGameName = GetSafeFileName(game.Name).ToLower();
                        if (!nameToId.ContainsKey(safeGameName))
                        {
                            nameToId.Add(safeGameName, new List<Guid>());
                        }
                        nameToId[safeGameName].Add(gameId);
                    }

                } else
                {
                    // Delete invalid shortcut
                    // System.IO.File.Delete(file);
                    logger.Warn(
                        $"Shortcut at {file} is not a valid shortcut created by ShortcutSync and" +
                        $" may cause it not to function as intended. If possible, " +
                        $"delete it or move it to another location.");
                }
            }
            return (games, nameToId);
        }


        /// <summary>
        /// Opens/Creates and returns a <see cref="IWshShortcut"/> at a 
        /// given path.
        /// </summary>
        /// <param name="shortcutPath">Full path to the shortcut.</param>
        /// <returns>The shortcut object.</returns>
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
                    logger.Error(ex, $"Could not open shortcut at \"{shortcutPath}\".");
                }
            }
            return null;
        }
    }
}