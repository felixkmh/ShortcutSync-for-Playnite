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
using Octokit;
using System.CodeDom.Compiler;
using System.Diagnostics;
using Microsoft.CSharp;
using System.Text;
using System.Drawing;
using System.Security.AccessControl;
using System.Windows.Forms;
using System.Drawing.IconLib;

namespace ShortcutSync
{
    public class ShortcutSync : Plugin
    {
        public static readonly ILogger logger = LogManager.GetLogger();
        private Thread thread;
        private ShortcutSyncSettings settings { get; set; }
        public ShortcutSyncSettingsView settingsView { get; set; }
        private Dictionary<Guid, string> existingShortcuts { get; set; }
        private Dictionary<string, IList<Guid>> shortcutNameToGameId { get; set; } = new Dictionary<string, IList<Guid>>();
        private Dictionary<Guid, Shortcut<Game>> Shortcuts { get; set; } = new Dictionary<Guid, Shortcut<Game>>();
        private ShortcutSyncSettings PreviousSettings { get; set; }

        public readonly Version version = new Version(1, 12);
        public override Guid Id { get; } = Guid.Parse("8e48a544-3c67-41f8-9aa0-465627380ec8");

        public enum UpdateStatus
        {
            None,
            Updated,
            Deleted,
            Created,
            Error
        }

        public class UpdateAvailableResult
        {
            public bool Available { get; private set; }
            public Version LatestVersion { get; private set; }
            public string Url { get; private set; }

            public UpdateAvailableResult(bool available, Version latestVersion, string url)
            {
                Available = available;
                LatestVersion = latestVersion;
                Url = url;
            }
        }


        public ShortcutSync(IPlayniteAPI api) : base(api)
        {
            settings = new ShortcutSyncSettings(this);
            TiledShortcut.FileDatabasePath = Path.Combine(api.Database.DatabasePath, "files");
        }

        public override IEnumerable<ExtensionFunction> GetFunctions()
        {
            return new List<ExtensionFunction>
            {
                new ExtensionFunction(
                    "Update All Shortcuts",
                    () =>
                    {
                        if (FolderIsAccessible(settings.ShortcutPath))
                        {
                            CreateFolderStructure(settings.ShortcutPath);
                            thread?.Join();
                            thread = new Thread(() =>
                            {
                                UpdateShortcutDicts(settings.ShortcutPath);
                                UpdateShortcuts(PlayniteApi.Database.Games);
                            });
                            thread.Start();
                        } else
                        {
                            PlayniteApi.Dialogs.ShowErrorMessage($"The selected shortcut folder \"{settings.ShortcutPath}\" is inaccessible. Please select another folder.", "Folder inaccessible.");
                        }

                    }),
                new ExtensionFunction(
                    "Create TiledShortcut for selected Games",
                    () =>
                    {
                        CreateShortcuts(PlayniteApi.MainView.SelectedGames);
                    }),
                new ExtensionFunction(
                    "Delete TiledShortcut for selected Games",
                    () =>
                        {
                            foreach (var game in PlayniteApi.MainView.SelectedGames)
                            {
                                if (Shortcuts.ContainsKey(game.Id))
                                {
                                    Shortcuts[game.Id].Remove();
                                    Shortcuts.Remove(game.Id);
                                }
                            }
                            foreach (var game in PlayniteApi.MainView.SelectedGames)
                            {
                                var name = game.Name.GetSafeFileName().ToLower();
                                if (shortcutNameToGameId.TryGetValue(name, out var copys))
                                {
                                    copys.Remove(game.Id);
                                    if (copys.Count == 0)
                                        shortcutNameToGameId.Remove(name);
                                    else if (copys.Count == 1)
                                        foreach (var id in copys)
                                        {
                                            var copy = PlayniteApi.Database.Games[id];
                                            if (copy != null && Shortcuts.ContainsKey(id))
                                                Shortcuts[id].Name = Path.GetFileNameWithoutExtension(GetShortcutPath(copy, false, settings.SeperateFolders));
                                        }
                                }

                            }
                    })
            };
        }

        public override void OnApplicationStarted()
        {
            UpdateShortcutDicts(settings.ShortcutPath);
            if (CreateFolderStructure(settings.ShortcutPath))
            {
                // (existingShortcuts, shortcutNameToGameId) = GetExistingShortcuts(settings.ShortcutPath);
                if (settings.UpdateOnStartup)
                {
                    if (FolderIsAccessible(settings.ShortcutPath))
                    {
                        thread?.Join();
                        thread = new Thread(() =>
                        {
                            UpdateShortcuts(PlayniteApi.Database.Games);
                        });
                        thread.Start();
                    }
                    else
                    {
                        PlayniteApi.Dialogs.ShowErrorMessage($"The selected shortcut folder \"{settings.ShortcutPath}\" is inaccessible. Please select another folder.", "Folder inaccessible.");
                    }
                } 
                PlayniteApi.Database.Games.ItemUpdated += Games_ItemUpdated;
                PlayniteApi.Database.Games.ItemCollectionChanged += Games_ItemCollectionChanged;
            } else
            {
                logger.Error("Could not create directory \"{settings.ShortcutPath}\". Try choosing another folder.");
            }
            settings.OnPathChanged += Settings_OnPathChanged;
            settings.OnSettingsChanged += Settings_OnSettingsChanged;
            PreviousSettings = settings.GetClone();
        }

        private void Settings_OnSettingsChanged()
        {
            thread?.Join();
            thread = new Thread(() =>
            {
                if (settings.ShortcutPath != PreviousSettings.ShortcutPath)
                {
                    UpdateShortcutDicts(PreviousSettings.ShortcutPath);
                    foreach (var shortcut in Shortcuts.Values) shortcut.Remove();
                }
                if (settings.SeperateFolders != PreviousSettings.SeperateFolders)
                {
                    UpdateShortcutDicts(settings.ShortcutPath);
                    foreach (var shortcut in Shortcuts.Values) shortcut.Remove();
                }
                UpdateShortcutDicts(settings.ShortcutPath);
                UpdateShortcuts(PlayniteApi.Database.Games);
                PreviousSettings = settings.GetClone();
            });
            thread.Start();
        }

        private void Settings_OnPathChanged(string oldPath, string newPath)
        {
            // if (MoveShortcutPath(oldPath, newPath))
            {
                if (CreateFolderStructure(newPath))
                {
                    UpdateShortcutDicts(newPath);
                }
            }
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
            CreateShortcuts(from game in e.AddedItems where ShouldKeepShortcut(game) select game);
            // Remove shortcuts to deleted games
            RemoveShortcuts(e.RemovedItems);
        }

        /// <summary>
        /// Callback handling game item updates. Updates the shortcuts 
        /// of all changed games.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Games_ItemUpdated(object sender, ItemUpdatedEventArgs<Game> e)
        {
            try
            {
                UpdateShortcuts(from update in e.UpdatedItems where SignificantChanges(update.OldData, update.NewData) select update.NewData, true);
            }
#pragma warning disable CS0168 // Variable ist deklariert, wird jedoch niemals verwendet
            catch (Exception ex)
#pragma warning restore CS0168 // Variable ist deklariert, wird jedoch niemals verwendet
            {
#if DEBUG
                logger.Debug(ex, $"Could not convert icon. Trying again...");
#endif
            }
        }

        private static bool SignificantChanges(Game oldData, Game newData) {
            if (oldData.Id              != newData.Id)              return true;
            if (oldData.Icon            != newData.Icon)            return true;
            if (oldData.BackgroundImage != newData.BackgroundImage) return true;
            if (oldData.CoverImage      != newData.CoverImage)      return true;
            if (oldData.Name            != newData.Name)            return true;
            if (oldData.Hidden          != newData.Hidden)          return true;
            if (oldData.Source          != newData.Source)          return true;
            return false;
        }

        public override void OnApplicationStopped()
        {
            // Unsubscribe from library change events.
            PlayniteApi.Database.Games.ItemUpdated -= Games_ItemUpdated;
            PlayniteApi.Database.Games.ItemCollectionChanged -= Games_ItemCollectionChanged;
            settings.OnSettingsChanged -= Settings_OnSettingsChanged;
            settings.OnPathChanged -= Settings_OnPathChanged;
        }


        public override ISettings GetSettings(bool firstRunSettings)
        {
            return settings;
        }

        public override System.Windows.Controls.UserControl GetSettingsView(bool firstRunSettings)
        {
            settingsView = new ShortcutSyncSettingsView();
            return settingsView;
        }

        public Task<UpdateAvailableResult> UpdateAvailable()
        {
            return Task.Run(() => {
            var github = new GitHubClient(new ProductHeaderValue("ShortcutSync-for-Playnite"));
            var remaining = github.GetLastApiInfo()?.RateLimit.Remaining;
            if (remaining == null)
            {
                try
                {
                    remaining = github.Miscellaneous.GetRateLimits().Result.Rate.Remaining;
                }
                catch (Exception)
                {
                    remaining = 0;
                }
            }
            if (remaining > 0)
            {
                try
                {
                    var release = github.Repository.Release.GetLatest("felixkmh", "ShortcutSync-for-Playnite").Result;
                    if (Version.TryParse(release.TagName.Replace("v",""), out Version latestVersion))
                    {
                        if (latestVersion > version)
                        {
                            logger.Info($"New version of ShortcutSync available. Current {version}, latest {latestVersion}");
                            return new UpdateAvailableResult(true, latestVersion, release.HtmlUrl);
                        }
                    }
                    logger.Debug($"Latest version: {release.TagName} parsed: {latestVersion}");
                }
#pragma warning disable CS0168 // Variable ist deklariert, wird jedoch niemals verwendet
                    catch (Exception ex)
#pragma warning restore CS0168 // Variable ist deklariert, wird jedoch niemals verwendet
                    {
#if DEBUG
                logger.Debug(ex, "Could not retrieve latest release.");
#endif
                }
            }
            return new UpdateAvailableResult(false, default, default);
            });
        }

        /// <summary>
        /// Update, create or remove a shortcut to a game.
        /// </summary>
        /// <param name="game">Game the shortcut should point to.</param>
        /// <param name="forceUpdate">If true, existing shortcut will be overwritten.</param>
        /// <param name="multiple">Specifies whether source name should be included in the shortcut
        /// name to prevent conflicts if multiple games have the same shortcut name.</param>
        /// <returns>Status of the associated shortcut.</returns>
        public UpdateStatus UpdateShortcut(Game game, bool forceUpdate, bool multiple = false, bool removed = false)
        {
            UpdateStatus status = UpdateStatus.None;
            // determine whether to create/update the shortcut or to delete it
            bool shortcutExists = existingShortcuts.TryGetValue(game.Id, out string currentPath);
            string desiredPath = GetShortcutPath(game, multiple);
            bool keepShortcut = ShouldKeepShortcut(game) && !removed;
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
            bool sourceEnabled = false;
            settings.SourceOptions.TryGetValue(GetSourceName(game), out sourceEnabled);
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
                if (Path.GetExtension(icon).ToLower() != ".ico")
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
            if (true || !game.IsInstalled || !settings.UsePlayAction || GetSourceName(game) == Constants.UNDEFINEDSOURCE)
            {
                CreateVbsLauncher(game, GetLauncherScriptPath(settings.ShortcutPath));
                CreateVisualElementsManifest(game, GetLauncherScriptPath(settings.ShortcutPath));
                CreateLnkURLToVbs(shortcutPath, icon, game);
            } 
            else 
            {
                if (GetSourceName(game) == "Xbox")
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
        public static bool ConvertToIcon(string icon, out string output)
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
        private string GetShortcutPath(Game game, bool includeSourceName = true, bool seperateFolders = false, string extension = ".lnk")
        {
            var validName = GetSafeFileName(game.Name);
            string path;

            if (seperateFolders)
            {
                path = Path.Combine(settings.ShortcutPath, GetSourceName(game));
            } else
            {
                path = Path.Combine(settings.ShortcutPath);
            }

            if (includeSourceName && !seperateFolders)
            {
                path = Path.Combine(path, validName + " (" + GetSourceName(game) + ")" + extension);
            } else
            {
                path = Path.Combine(path, validName + extension);
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
        public void UpdateShortcuts_Alt(IEnumerable<Game> gamesToUpdate, bool forceUpdate = false, bool removed = false)
        {
            bool errorOccured = false;
            {
                {
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
                        if (ShouldKeepShortcut(game) && !removed)
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
                                            logger.Error(ex, $"Could not move {currentPath} to {newPath}, while updating {game.Name}.");
                                        }
                                    }
                                }
                            }
                            if (UpdateShortcut(game, forceUpdate, existing.Count > 0, removed) == UpdateStatus.Error)
                            {
                                errorOccured = true;
                            }
                        } else if (existing.Count == 1)
                        {
                            if (UpdateShortcut(game, forceUpdate, false, removed) == UpdateStatus.Error)
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
                                        logger.Error(ex, $"Could not move {currentPath} to {newPath}, while updating {game.Name}.");
                                    }
                                    existingShortcuts[existing[0].Id] = newPath;
                                }
                            }
                        } else
                        {
                            if (UpdateShortcut(game, forceUpdate, false, removed) == UpdateStatus.Error)
                            {
                                errorOccured = true;
                            }
                        }
                      
                    }
                    if (errorOccured)
                    {
                        // Refresh dictionaries to properly reflect current state.
                        // (existingShortcuts, shortcutNameToGameId) = GetExistingShortcuts(settings.ShortcutPath);
                        UpdateShortcutDicts(settings.ShortcutPath);
                    }

                }

            }

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
                description: "Launch " + game.Name + " on " + GetSourceName(game) + "." + $" [{game.Id}]",
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
                description: "Launch " + game.Name + " on " + GetSourceName(game) + "." + $" [{game.Id}]",
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
                description: "Launch " + game.Name + " on " + GetSourceName(game) + " via Playnite." + $" [{game.Id}]");
        }

        public void CreateLnkURLToVbs(string path, string iconPath, Game game)
        {
            CreateLnk(
                shortcutPath: path,
                targetPath: $"{GetLauncherScriptPath(settings.ShortcutPath)}\\{game.Id}.vbs",
                iconPath: iconPath,
                description: "Launch " + game.Name + " on " + GetSourceName(game) + " via Playnite." + $" [{game.Id}]");
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
                description: "Launch " + game.Name + " on " + GetSourceName(game) + "." + $" [{game.Id}]");
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
                if (shortcut != null)
                {
                    shortcut.IconLocation = iconPath;
                    shortcut.TargetPath = targetPath;
                    shortcut.Description = description;
                    shortcut.WorkingDirectory = workingDirectory;
                    shortcut.Arguments = arguments;
                    shortcut.Save();
                } else
                {
                    logger.Error($"Could not create shortcut at \"{shortcutPath}\".");
                }
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
                        logger.Warn($"Shortcut at \"{file}\" conatins Guid \"{gameId}\", but this Id is not contained in the game database. Will be deleted.");
                        try
                        {
                            System.IO.File.Delete(file);
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex, $"Could not remove shortcut at \"{file}\".");
                        }
                    }
                    else if (games.ContainsKey(game.Id))
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

        public void UpdateShortcuts(IEnumerable<Game> games, bool forceUpdate = false)
        {
            CreateShortcuts(from game in games where ShouldKeepShortcut(game) select game);
            RemoveShortcuts(from game in games where !ShouldKeepShortcut(game) select game);
            if (forceUpdate) foreach(var game in games) if (Shortcuts.ContainsKey(game.Id)) Shortcuts[game.Id].Update(true); 
        }

        public void CreateShortcuts(IEnumerable<Game> games)
        {
            foreach (var game in games)
            {
                if (Shortcuts.ContainsKey(game.Id))
                {
                    if (!Shortcuts[game.Id].Exists)
                    {
                        Shortcuts[game.Id].Remove();
                        Shortcuts.Remove(game.Id);
                        if (shortcutNameToGameId.ContainsKey(game.Name.GetSafeFileName().ToLower()))
                            shortcutNameToGameId[game.Name.GetSafeFileName().ToLower()].Remove(game.Id);
                    }
                }
            }
            foreach (var game in games)
            {
                var hasDuplicates = false;
                if (shortcutNameToGameId.TryGetValue(game.Name.GetSafeFileName().ToLower(), out var copyIds))
                {
                    foreach (var copyId in copyIds)
                    {
                        var copy = PlayniteApi.Database.Games[copyId];
                        if (copy != null && copyId != game.Id)
                        {
                            Shortcuts[copyId].Name = Path.GetFileNameWithoutExtension(GetShortcutPath(copy, true, settings.SeperateFolders));
                            hasDuplicates = true;
                        }
                    }
                    shortcutNameToGameId[game.Name.GetSafeFileName().ToLower()].AddMissing(game.Id);
                }
                else
                {
                    shortcutNameToGameId[game.Name.GetSafeFileName().ToLower()] = new List<Guid>() { game.Id };
                }
                if (Shortcuts.TryGetValue(game.Id, out Shortcut<Game> existing))
                {
                    existing.Name = Path.GetFileNameWithoutExtension(GetShortcutPath(game, hasDuplicates, settings.SeperateFolders));
                } else
                {
                    if (settings.UsePlayAction)
                    {
                        Shortcuts.Add(game.Id,
                            new TiledShortcutsPlayAction
                            (
                                targetGame: game,
                                shortcutPath: GetShortcutPath(game, hasDuplicates, settings.SeperateFolders),
                                launchScriptFolder: GetLauncherScriptPath(settings.ShortcutPath),
                                tileIconFolder: GetLauncherScriptIconsPath(settings.ShortcutPath)
                            )
                        );
                    } else
                    {
                        Shortcuts.Add(game.Id,
                            new TiledShortcut
                            (
                                targetGame: game,
                                shortcutPath: GetShortcutPath(game, hasDuplicates, settings.SeperateFolders),
                                launchScriptFolder: GetLauncherScriptPath(settings.ShortcutPath),
                                tileIconFolder: GetLauncherScriptIconsPath(settings.ShortcutPath)
                            )
                        );
                    }
                }
            }
            Parallel.ForEach(games, new ParallelOptions() { MaxDegreeOfParallelism = Math.Max(Environment.ProcessorCount / 3, 1) }, game => Shortcuts[game.Id]?.CreateOrUpdate());
        }

        public void RemoveShortcuts(IEnumerable<Game> games)
        {
            foreach (var game in games)
            {
                if (Shortcuts.ContainsKey(game.Id))
                {
                    Shortcuts[game.Id].Remove();
                    Shortcuts.Remove(game.Id);
                }
            }
            foreach (var game in games)
            {
                var name = game.Name.GetSafeFileName().ToLower();
                if (shortcutNameToGameId.TryGetValue(name, out var copys))
                {
                    copys.Remove(game.Id);
                    if (copys.Count == 0)
                        shortcutNameToGameId.Remove(name);
                    else if (copys.Count == 1)
                        foreach (var id in copys)
                        {
                            var copy = PlayniteApi.Database.Games[id];
                            if (copy != null && Shortcuts.ContainsKey(id))
                                Shortcuts[id].Name = Path.GetFileNameWithoutExtension(GetShortcutPath(copy, false, settings.SeperateFolders));
                        }
                }

            }
        }

        public void RemoveFromShortcutDicts(Guid gameId)
        {
            if (Shortcuts.TryGetValue(gameId, out var shortcut))
            {
                shortcut.Remove();
                Shortcuts.Remove(gameId);
            }

            var game = PlayniteApi.Database.Games.Get(gameId);
            if (game != null)
            {
                var name = game.Name.GetSafeFileName().ToLower();
                if (shortcutNameToGameId.TryGetValue(name, out var copys))
                {
                    copys.Remove(gameId);
                    if (copys.Count == 0) shortcutNameToGameId.Remove(name);
                }
            }
        }

        public void UpdateShortcutDicts(string folderPath, string shortcutName = "")
        {
            Shortcuts.Clear();
            shortcutNameToGameId.Clear();
            var files = Directory.GetFiles(folderPath, shortcutName + "*.lnk", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                var lnk = OpenLnk(file);
                if (lnk != null && ExtractIdFromLnkDescription(lnk.Description, out Guid gameId))
                {
                    var game = PlayniteApi.Database.Games.Get(gameId);
                    if (game != null)
                    {
                        if (Shortcuts.TryGetValue(game.Id, out var shortcut))
                        {
                            shortcut.Remove();
                        }
                        if (false && settings.UsePlayAction)
                        {
                            string workingDirectory = game.PlayAction.WorkingDir;
                            string targetPath = game.PlayAction.Path;

                            if (game.PlayAction.Type != GameActionType.URL)
                            {
                                workingDirectory = PlayniteApi.ExpandGameVariables(game, game.PlayAction.WorkingDir);
                                targetPath = PlayniteApi.ExpandGameVariables(game, game.PlayAction.Path);
                                if (!string.IsNullOrEmpty(workingDirectory))
                                {
                                    targetPath = Path.Combine(workingDirectory, targetPath);
                                }
                            } 
                            Shortcuts[game.Id] = new TiledShortcutsPlayAction(game, file, GetLauncherScriptPath(settings.ShortcutPath), GetLauncherScriptIconsPath(settings.ShortcutPath),
                                workingDirectory, targetPath, game.PlayAction.Arguments?? "");
                        }
                        else
                            Shortcuts[game.Id] = new TiledShortcut(game, file, GetLauncherScriptPath(settings.ShortcutPath), GetLauncherScriptIconsPath(settings.ShortcutPath));
                        string safeGameName = game.Name.GetSafeFileName().ToLower();
                        if (!shortcutNameToGameId.ContainsKey(safeGameName))
                        {
                            shortcutNameToGameId.Add(safeGameName, new List<Guid>());
                        }
                        shortcutNameToGameId[safeGameName].Add(gameId);
                    }
                }
            }
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

        /// <summary>
        /// Checks whether a folder can be written to.
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns>Whether the folder can be written to.</returns>
        private bool FolderIsAccessible(string folderPath)
        {
            bool accessible = false;
            try
            {
                if (Directory.Exists(folderPath))
                {
                    var ac = Directory.GetAccessControl(folderPath);
                    accessible = true;
                }
            }
            catch (UnauthorizedAccessException)
            {
            }
            return accessible;
        }

        /// <summary>
        /// Safely gets a name of the source for a game.
        /// </summary>
        /// <param name="game"></param>
        /// <returns>Source name or "Undefined" if no source exists.</returns>
        private string GetSourceName(Game game)
        {
            if (game.Source == null)
            {
                return Constants.UNDEFINEDSOURCE;
            } else
            {
                return game.Source.Name;
            }
        }

        private bool CreateExecutableShortcut(Game game, string folderPath)
        {
            CSharpCodeProvider codeProvider = new CSharpCodeProvider();
            string output = Path.Combine(folderPath, $"{game.Id}.exe");
            StringBuilder compilerOptions = new StringBuilder();
            string iconPath = Path.Combine(PlayniteApi.Database.GetFileStoragePath(game.Id), Path.GetFileName(game.Icon));
            compilerOptions.Append("/optimize ");
            if (Path.GetExtension(iconPath).ToLower() == ".ico")
            {
                compilerOptions.Append($"/win32icon:\"");
                compilerOptions.Append(iconPath);
                compilerOptions.Append("\"");
            }
            var parameters = new CompilerParameters
            {
                GenerateExecutable = true,
                OutputAssembly = output,
                CompilerOptions = compilerOptions.ToString()
            };
            parameters.ReferencedAssemblies.Add("System.dll");
            parameters.ReferencedAssemblies.Add("System.Diagnostics.Process.dll");
            var results = codeProvider.CompileAssemblyFromSource(parameters, 
                "using System;\n" +
                "using System.Diagnostics;\n" +
                "namespace ShortcutLauncher{\n" +
                    "class Launch{\n" +
                        "static void Main(string[] args){\n" +
                            "var startInfo = new ProcessStartInfo();\n" +
                            $"startInfo.FileName= \"cmd.exe\";\n" +
                            $"startInfo.Arguments = \"/c start playnite://playnite/start/{game.Id}\";\n" +
                            "startInfo.CreateNoWindow = true;\n" +
                            "startInfo.UseShellExecute = false;\n" +
                            "startInfo.WindowStyle = ProcessWindowStyle.Hidden;\n" +
                            "System.Diagnostics.Process.Start(startInfo);\n" +
                        "}\n" +
                    "}\n" +
                "}");

            if (results.Errors.Count > 0)
            {
                string errors = "";
                foreach (CompilerError CompErr in results.Errors)
                {
                    errors = errors +
                                "Line number " + CompErr.Line +
                                ", Error Number: " + CompErr.ErrorNumber +
                                ", '" + CompErr.ErrorText + ";" +
                                Environment.NewLine + Environment.NewLine;
                }
                logger.Error($"Could not compile launcher for {game.Name}. Errors:\n" + errors);
                return false;
            } else
            {
                logger.Debug($"Succesfully compiled launcher for {game.Name}");
                return true;
            }
        }

        private string CreateVbsLauncher(Game game, string folderPath)
        {
            string fileName = $"{game.Id}.vbs";
            string fullPath = Path.Combine(folderPath, fileName);
            string script = 
                "Dim prefix, id\n" +

                "prefix = \"playnite://playnite/start/\"\n" +
                $"id = \"{game.Id}\"\n" +

                "Set WshShell = WScript.CreateObject(\"WScript.Shell\")\n" +
                "WshShell.Run prefix &id, 1";
            try
            {
                using (var scriptFile = System.IO.File.CreateText(fullPath))
                {
                    scriptFile.Write(script);
                }
                return fullPath;
            }
            catch (Exception)
            {

            }
            
            return string.Empty;
        }

        private bool CreateLauncherShortcut(Game game, string shortcutFolderPath, string launcherFolderPath)
        {
            CreateLnk(
                shortcutPath: Path.Combine(shortcutFolderPath, game.Name + ".lnk"),
                targetPath: Path.Combine(launcherFolderPath, game.GameId + ".vbs"),
                iconPath: "",
                description: ""
                );
            return true;
        }

        private string CreateVisualElementsManifest(Game game, string folderPath)
        {
#if DEBUG
            logger.Debug($"Creating xml for {game.Name}.");
#endif
            string fileName = $"{game.Id}.visualelementsmanifest.xml";
            string fullPath = Path.Combine(folderPath, fileName);
            string foregroundTextStyle = "light";
            string backgroundColorCode = "#000000";
            if (   !Path.GetFileName(game.Icon).IsNullOrEmpty()
                && System.IO.File.Exists(GetIconPath(game)))
            {
                string iconPath = GetIconPath(game);
                Bitmap bitmap = null;
                if (Path.GetExtension(game.Icon).ToLower() == ".ico")
                {
                    using (var stream = System.IO.File.OpenRead(iconPath))
                    {
                        var i = new System.Drawing.IconLib.MultiIcon();
                        i.Load(stream);
                        int maxIndex = 0;
                        int maxWidth = 0;
                        int index = 0;
                        foreach(var imag in i[0])
                        {
                            if (imag.Size.Width >= 150 && imag.Size.Height >= 150)
                            {
                                // logger.Info($"Icon size: {imag.Size} for {game.Name}.");
                                bitmap = imag.Icon.ToBitmap();
                                break;
                            }
                            if (imag.Size.Width > maxWidth)
                            {
                                maxIndex = index;
                                maxWidth = imag.Size.Width;
                                bitmap = imag.Icon.ToBitmap();
                            }
                            ++index;
                        }
                        

                        using (var icon = new Icon(iconPath, 1000, 1000))
                        {
                            // logger.Info($"Icon size: {icon.Size} for {game.Name}.");
                            // bitmap = icon.ToBitmap();
                        }
                    }
                } else
                {
                    bitmap = new Bitmap(iconPath);
                }
                if (bitmap != null && bitmap.Width > 0 && bitmap.Height > 0)
                {
                    var brightness = GetAverageBrightness(bitmap);
                    foregroundTextStyle = brightness > 0.4f ? "dark" : "light";
                    // var colorThief = new ColorThiefDotNet.ColorThief();
                    // backgroundColorCode = colorThief.GetColor(bitmap, 10, false).Color.ToHexString();
                    var bgColor = GetDominantColor(bitmap, brightness);
                    backgroundColorCode = $"#{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";

                    // resize
                    int newWidth = 150;
                    int newHeight = 150;
                    if (bitmap.Width >= bitmap.Height)
                    {
                        float scale = (float)newHeight / bitmap.Height;
                        newWidth = (int)Math.Round(bitmap.Width * scale);
                    } else
                    {
                        float scale = (float)newWidth / bitmap.Width;
                        newHeight = (int)Math.Round(bitmap.Height * scale);
                    }
                    Bitmap resized = new Bitmap(newWidth, newHeight);
                    using (Graphics graphics = Graphics.FromImage(resized))
                    {
                        if (bitmap.Width <= 64 && bitmap.Height <= 64)
                        {
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                        }
                        graphics.DrawImage(bitmap, 0, 0, newWidth, newHeight);
                    }

                    resized.Save(Path.Combine(folderPath, Constants.ICONFOLDERNAME, $"{game.Id}.png"), ImageFormat.Png);
                    resized.Dispose();
                }
                bitmap.Dispose();
            }

            string script =
                "<Application xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n" + 
                "<VisualElements\n" + 
                $"BackgroundColor = \"{backgroundColorCode}\"\n" + 
                "ShowNameOnSquare150x150Logo = \"on\"\n" + 
                $"ForegroundText = \"{foregroundTextStyle}\"\n" +
                $"Square150x150Logo = \"{Constants.ICONFOLDERNAME}\\{game.Id}.png\"\n" + 
                $"Square70x70Logo = \"{Constants.ICONFOLDERNAME}\\{game.Id}.png\"/>\n" + 
                "</Application>";
            try
            {
                using (var scriptFile = System.IO.File.CreateText(fullPath))
                {
                    scriptFile.Write(script);
                }
                return fullPath;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Could not create or write to launcher script \"{fullPath}\".");
            }

            return string.Empty;
        }

        private float GetAverageBrightness(Bitmap bitmap)
        {
            float accumulator = 0F;
            int count = 0;
            for (int y = 0; y < bitmap.Height; ++y)
                for (int x = 0; x < bitmap.Width; ++x)
                {
                    var pixelColor = bitmap.GetPixel(x, y);
                    if (pixelColor.A > 0)
                    {
                        accumulator += pixelColor.GetBrightness();
                        ++count;
                    }
                }
            return count > 0 ? accumulator / count : 0f;
        }

        private Color GetDominantColor(Bitmap bitmap, float brightness = 0.5f)
        {
            float r = 0f, g = 0f, b = 0f;
            for (int y = 0; y < bitmap.Height; ++y)
                for (int x = 0; x < bitmap.Width; ++x)
                {
                    var pixelColor = bitmap.GetPixel(x, y);
                    
                    float alpha = pixelColor.A / 255.0f;
                    r += (alpha / 255.0f) * pixelColor.R;
                    g += (alpha / 255.0f) * pixelColor.G;
                    b += (alpha / 255.0f) * pixelColor.B;
                }
            r *= r;
            g *= g;
            b *= b;
            var max = Math.Max(r, Math.Max(g, b));
            max = max <= 0.00001 ? 1 : max;
            r = 255 * r / max;
            g = 255 * g / max;
            b = 255 * b / max;
            var color = System.Drawing.Color.FromArgb(
                    alpha: 255, 
                    red:   Math.Max(0, (int)Math.Min(255, r)), 
                    green: Math.Max(0, (int)Math.Min(255, g)), 
                    blue:  Math.Max(0, (int)Math.Min(255, b))
                );
            var brightnessFactor = 0;
            if (Math.Abs(brightness) - Math.Abs(color.GetBrightness()) < 0.3)
                brightnessFactor = brightness >= 0.5f ? -1 : 1;
            color = System.Drawing.Color.FromArgb(
                    alpha: 255,
                    red: Math.Max(0, (int)Math.Min(255, color.R + brightnessFactor * 50)),
                    green: Math.Max(0, (int)Math.Min(255, color.G + brightnessFactor * 50)),
                    blue: Math.Max(0, (int)Math.Min(255, color.B + brightnessFactor * 50))
                );
            return color;
        }

        private Color GetDominantColorQuantized(Bitmap bitmap, int transparencyThreshold = 10)
        {
            Dictionary<int, int> colors = new Dictionary<int, int>();
            for (int y = 0; y < bitmap.Height; ++y)
                for (int x = 0; x < bitmap.Width; ++x)
                {
                    var pixelColor = bitmap.GetPixel(x, y);
                    if (pixelColor.A > transparencyThreshold)
                    {
                        var closestColor = GetClosestColor(pixelColor).ToArgb();
                        if (colors.TryGetValue(closestColor, out int count))
                        {
                            colors[closestColor] = count + 1;
                        } else
                        {
                            colors.Add(closestColor, 1);
                        }
                    }
                }
            if (colors.Count > 0)
            {
                return Color.FromArgb(colors.Aggregate((l, r) => l.Value > r.Value ? l : r).Key);
            }
            else
            {
                return Color.Black;
            }
        }

        private Color GetClosestColor(Color color)
        {
            float minDist = float.PositiveInfinity;
            KnownColor closestColor = KnownColor.White;

            float sqrDistance(Color a, Color b)
            {
                return (a.R - b.R) * (a.R - b.R) + (a.G - b.G) * (a.G - b.G) + (a.B - b.B) * (a.B - b.B);
            }
            foreach (KnownColor knownColor in Enum.GetValues(typeof(KnownColor)))
            {
                var dist = sqrDistance(Color.FromKnownColor(knownColor), color);
                if (dist < minDist)
                {
                    minDist = dist;
                    closestColor = knownColor;
                }
            }
            return Color.FromKnownColor(closestColor);
        }

        private string GetLauncherScriptPath(string shortcutPath)
        {
            return Path.Combine(shortcutPath, Constants.LAUNCHSCRIPTFOLDERNAME);
        }

        private string GetLauncherScriptIconsPath(string shortcutPath)
        {
            return Path.Combine(shortcutPath, Constants.LAUNCHSCRIPTFOLDERNAME, Constants.ICONFOLDERNAME);
        }

        public bool MoveShortcutPath(string oldPath, string newPath) {
            try
            {
                if (Directory.Exists(oldPath))
                {
                    MergeMoveDirectory(oldPath, newPath, true);
                }
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Could not move folder \"{oldPath}\" to \"{newPath}\".");
                return false;
            }
        }

        private bool CreateFolderStructure(string path)
        {
            try
            {
                Directory.CreateDirectory(path);
                var launchScriptsFolder = Directory.CreateDirectory(GetLauncherScriptPath(path));
                Directory.CreateDirectory(GetLauncherScriptIconsPath(path));
                launchScriptsFolder.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Could not create folder structure at \"{path}\".");
                return false;
            }
        } 

        private bool MergeMoveDirectory(string source, string target, bool overwrite = false)
        {
            if (Directory.Exists(target))
            {
                foreach(var dir in Directory.GetDirectories(source))
                {
                    MergeMoveDirectory(dir, Path.Combine(target, Path.GetFileName(dir)));
                }
                foreach(var file in Directory.GetFiles(source))
                {
                    var newFile = Path.Combine(target, Path.GetFileName(file));
                    if (!System.IO.File.Exists(newFile))
                    {
                        System.IO.File.Move(file, newFile);
                    } else if (overwrite)
                    {
                        System.IO.File.Delete(newFile);
                        System.IO.File.Move(file, newFile);
                    }
                }
                Directory.Delete(source, true);
            } else
            {
                Directory.Move(source, target);
            }
            return true;
        }

        private string GetIconPath(Game game)
        {
            return Path.Combine(PlayniteApi.Database.GetFileStoragePath(game.Id), Path.GetFileName(game.Icon));
        }
    }
}