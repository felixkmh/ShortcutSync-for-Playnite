using Microsoft.CSharp;
using Octokit;
using Playnite.SDK;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using System;
using System.Reflection;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ShortcutSync
{
    public class ShortcutSync : Plugin
    {
        public static readonly ILogger logger = LogManager.GetLogger();
        private Task backgroundTask = Task.CompletedTask;
        private ShortcutSyncSettings settings { get; set; }
        public ShortcutSyncSettingsView settingsView { get; set; }
        private Dictionary<string, IList<Guid>> shortcutNameToGameId { get; set; } = new Dictionary<string, IList<Guid>>();
        private Dictionary<Guid, Shortcut<Game>> existingShortcuts { get; set; } = new Dictionary<Guid, Shortcut<Game>>();
        private ShortcutSyncSettings previousSettings { get; set; }

        public readonly Version version = new Version(1, 14);
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

#if VERSION7
        public override IEnumerable<ExtensionFunction> GetFunctions()
        {
            return new List<ExtensionFunction>
            {
                new ExtensionFunction(
                    "Update All Shortcuts",
                    () =>
                    {
                        CreateFolderStructure(settings.ShortcutPath);
                        if (FolderIsAccessible(settings.ShortcutPath))
                        {
                            CreateFolderStructure(settings.ShortcutPath);
                            backgroundTask = backgroundTask.ContinueWith((_) => {
                                UpdateShortcutDicts(settings.ShortcutPath, settings.Copy());
                                UpdateShortcuts(PlayniteApi.Database.Games, settings.Copy());
                            });
                        } else
                        {
                            PlayniteApi.Dialogs.ShowErrorMessage($"The selected shortcut folder \"{settings.ShortcutPath}\" is inaccessible. Please select another folder.", "Folder inaccessible.");
                        }

                    }),
                new ExtensionFunction(
                    "Create manual TiledShortcut for selected Games",
                    () =>
                    {
                        AddShortcutsManually(from game in PlayniteApi.MainView.SelectedGames select game.Id, settings);
                    }),
                new ExtensionFunction(
                    "Delete manual TiledShortcut for selected Games",
                    () =>
                    {
                        RemoveShortcutsManually(from game in PlayniteApi.MainView.SelectedGames select game.Id, settings);    
                    }),
                new ExtensionFunction(
                    "Exclude selected games from ShortcutSync",
                    () =>
                    {
                        AddToExclusionList(from game in PlayniteApi.MainView.SelectedGames select game.Id, settings);
                    }),
                new ExtensionFunction(
                    "Remove selected Games from ShortcutSync exclusion list",
                    () =>
                    {
                        RemoveFromExclusionList(from game in PlayniteApi.MainView.SelectedGames select game.Id, settings);
                    })
            };
        }
#endif

        public void RemoveFromExclusionList(IEnumerable<Guid> gameIds, ShortcutSyncSettings settings)
        {
            foreach (var id in gameIds)
                settings.ExcludedGames.Remove(id);
            UpdateShortcuts(from id in gameIds where PlayniteApi.Database.Games.Get(id) != null select PlayniteApi.Database.Games.Get(id), settings.Copy());
            SavePluginSettings(settings);
        }

        public void AddToExclusionList(IEnumerable<Guid> gameIds, ShortcutSyncSettings settings)
        {
            foreach (var id in gameIds)
                settings.ExcludedGames.Add(id);
            UpdateShortcuts(from id in gameIds where PlayniteApi.Database.Games.Get(id) != null select PlayniteApi.Database.Games.Get(id), settings.Copy());
            SavePluginSettings(settings);
        }

#if VERSION8
        public override List<GameMenuItem> GetGameMenuItems(GetGameMenuItemsArgs args)
        {
            return new List<GameMenuItem>
            {
                new GameMenuItem
                {
                    Description = "Create manual TiledShortcut(s)",
                    MenuSection = "ShortcutSync",
                    Action = context => 
                    { 
                        backgroundTask = backgroundTask.ContinueWith((_) => AddShortcutsManually(from game in context.Games select game.Id, settings)); 
                    }
                },
                new GameMenuItem
                {
                    Description = "Remove manual TiledShortcut(s)",
                    MenuSection = "ShortcutSync",
                    Action = context => 
                    {
                        backgroundTask = backgroundTask.ContinueWith((_) => RemoveShortcutsManually(from game in context.Games select game.Id, settings)); 
                    }
                },
                new GameMenuItem
                {
                    Description = "Update TiledShortcut(s)",
                    MenuSection = "ShortcutSync",
                    Action = context => {
                        backgroundTask = backgroundTask.ContinueWith((_) => UpdateShortcuts(context.Games, settings.Copy(), true)); 
                    }
                },
                new GameMenuItem
                {
                    Description = "Exclude selected games from ShortcutSync",
                    MenuSection = "ShortcutSync",
                    Action = context => {
                        AddToExclusionList(from game in context.Games select game.Id, settings);
                    }
                },
                new GameMenuItem
                {
                    Description = "Remove selected Games from ShortcutSync exclusion list",
                    MenuSection = "ShortcutSync",
                    Action = context => {
                        RemoveFromExclusionList(from game in context.Games select game.Id, settings);
                    }
                }
            };
        }

        public override List<MainMenuItem> GetMainMenuItems(GetMainMenuItemsArgs args)
        {
            return new List<MainMenuItem>
            {
                new MainMenuItem
                {
                    Description = "Update All Shortcuts",
                    MenuSection = "@|ShortcutSync",
                    Action = context => {
                        PlayniteApi.Dialogs.ActivateGlobalProgress(
                            (progress) => { 
                                UpdateShortcutDicts(settings.ShortcutPath, settings.Copy());
                                UpdateShortcuts(PlayniteApi.Database.Games, settings.Copy()); 
                            }, 
                            new GlobalProgressOptions("Updating Shortcuts...", false)
                        );
                        CreateFolderStructure(settings.ShortcutPath);
                        backgroundTask = backgroundTask.ContinueWith((_) => {
                            UpdateShortcutDicts(settings.ShortcutPath, settings.Copy());
                            UpdateShortcuts(PlayniteApi.Database.Games, settings.Copy());
                        });
                    }
                }
            };
        }
#endif
        public override void OnApplicationStarted()
        {
            CreateFolderStructure(settings.ShortcutPath);
            UpdateShortcutDicts(settings.ShortcutPath,settings);
            if (CreateFolderStructure(settings.ShortcutPath))
            {
                // (existingShortcuts, shortcutNameToGameId) = GetExistingShortcuts(settings.ShortcutPath);
                if (settings.UpdateOnStartup)
                {
                    if (FolderIsAccessible(settings.ShortcutPath))
                    {
                        backgroundTask = backgroundTask.ContinueWith((_) =>
                        {
                            UpdateShortcuts(PlayniteApi.Database.Games, settings.Copy());
                        });
                    }
                    else
                    {
                        PlayniteApi.Dialogs.ShowErrorMessage($"The selected shortcut folder \"{settings.ShortcutPath}\" is inaccessible. Please select another folder.", "Folder inaccessible.");
                    }
                }
                PlayniteApi.Database.Games.ItemUpdated += Games_ItemUpdated;
                PlayniteApi.Database.Games.ItemCollectionChanged += Games_ItemCollectionChanged;
            }
            else
            {
                logger.Error("Could not create directory \"{settings.ShortcutPath}\". Try choosing another folder.");
            }
            settings.OnPathChanged += Settings_OnPathChanged;
            settings.OnSettingsChanged += Settings_OnSettingsChanged;
            previousSettings = settings.Copy();
        }

        private void Settings_OnSettingsChanged()
        {
            backgroundTask = backgroundTask.ContinueWith((_) =>
            {
                var settingsSnapshot = settings.Copy();
                foreach(var playActionOpt in settings.EnabledPlayActions)
                {
                    if (playActionOpt.Value != previousSettings.EnabledPlayActions[playActionOpt.Key])
                        foreach (var shortcut in existingShortcuts.Values) {
                            if (GetSourceName(shortcut.TargetObject) == playActionOpt.Key)
                                shortcut.Remove();
                        }
                }
                if (settingsSnapshot.ShortcutPath != previousSettings.ShortcutPath
                || settingsSnapshot.SeparateFolders != previousSettings.SeparateFolders)
                {
                    UpdateShortcutDicts(previousSettings.ShortcutPath, settingsSnapshot);
                    foreach (var shortcut in existingShortcuts.Values)
                    {
                        bool moved = shortcut.Move(
                            GetShortcutPath(
                                shortcut.TargetObject,
                                settingsSnapshot.ShortcutPath,
                                HasExistingShortcutDuplicates(shortcut.TargetObject) && !settingsSnapshot.SeparateFolders,
                                settingsSnapshot.SeparateFolders),
                            GetLauncherScriptIconsPath(),
                            GetLauncherScriptPath()
                        );
                        if (!moved) shortcut.Remove();
                    }
                }
                UpdateShortcutDicts(settingsSnapshot.ShortcutPath, settingsSnapshot);
                UpdateShortcuts(PlayniteApi.Database.Games, settingsSnapshot);
                previousSettings = settingsSnapshot;
            });
        }

        private void Settings_OnPathChanged(string oldPath, string newPath)
        {
            // if (MoveShortcutPath(oldPath, newPath))
            {
                if (CreateFolderStructure(newPath))
                {
                    UpdateShortcutDicts(newPath, settings.Copy());
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
            CreateShortcuts(from game in e.AddedItems where ShouldKeepShortcut(game, settings.Copy()) select game, settings.Copy());
            // Remove shortcuts to deleted games
            RemoveShortcuts(e.RemovedItems, settings.Copy());
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
                UpdateShortcuts(from update in e.UpdatedItems where SignificantChanges(update.OldData, update.NewData) select update.NewData, settings.Copy(), true);
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

        private static bool SignificantChanges(Game oldData, Game newData)
        {
            if (oldData.Id != newData.Id) return true;
            if (oldData.Icon != newData.Icon) return true;
            if (oldData.BackgroundImage != newData.BackgroundImage) return true;
            if (oldData.CoverImage != newData.CoverImage) return true;
            if (oldData.Name != newData.Name) return true;
            if (oldData.Hidden != newData.Hidden) return true;
            if (oldData.Source != newData.Source) return true;
            if (oldData.IsInstalled != newData.IsInstalled) return true;
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
            return Task.Run(() =>
            {
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
                        if (Version.TryParse(release.TagName.Replace("v", ""), out Version latestVersion))
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

        public void AddShortcutsManually(IEnumerable<Guid> gameIds, ShortcutSyncSettings settings)
        {
            foreach (var id in gameIds)
                settings.ManuallyCreatedShortcuts.Add(id);
            UpdateShortcuts(from id in gameIds where PlayniteApi.Database.Games.Get(id) != null select PlayniteApi.Database.Games.Get(id), settings.Copy());
            SavePluginSettings(settings);
        }

        public void RemoveShortcutsManually(IEnumerable<Guid> gameIds, ShortcutSyncSettings settings)
        {
            foreach (var id in gameIds)
                settings.ManuallyCreatedShortcuts.Remove(id);
            UpdateShortcuts(from id in gameIds where PlayniteApi.Database.Games.Get(id) != null select PlayniteApi.Database.Games.Get(id), settings.Copy());
            SavePluginSettings(settings);
        }

        /// <summary>
        /// Checks whether a game's shortcut will be kept when updated
        /// based on its state and the plugin settings.
        /// </summary>
        /// <param name="game">The game that is ckecked.</param>
        /// <returns>Whether shortcut should be kept.</returns>
        private static bool ShouldKeepShortcut(Game game, ShortcutSyncSettings settings)
        {
            if (game == null) return false;
            bool sourceEnabled = false;
            settings.SourceOptions.TryGetValue(GetSourceName(game), out sourceEnabled);
            bool exludeBecauseHidden = settings.ExcludeHidden && game.Hidden;
            bool keepShortcut =
                (game.IsInstalled ||
                !settings.InstalledOnly) &&
                sourceEnabled &&
                !exludeBecauseHidden;
            return !settings.ExcludedGames.Contains(game.Id) && (keepShortcut || settings.ManuallyCreatedShortcuts.Contains(game.Id));
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
        private static string GetShortcutPath(Game game, string basePath, bool includeSourceName = true, bool seperateFolders = false, string extension = ".lnk")
        {
            var validName = GetSafeFileName(game.Name);
            string path;

            if (seperateFolders)
            {
                path = Path.Combine(basePath, GetSourceName(game));
            }
            else
            {
                path = Path.Combine(basePath);
            }

            if (includeSourceName && !seperateFolders)
            {
                path = Path.Combine(path, validName + " (" + GetSourceName(game) + ")" + extension);
            }
            else
            {
                path = Path.Combine(path, validName + extension);
            }
            return path;
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
        /// Checks whether a game entry changes because it
        /// was launched or closed.
        /// </summary>
        /// <param name="data">The event data of the changed game.</param>
        /// <returns></returns>
        private static bool WasLaunchedOrClosed(ItemUpdateEvent<Game> data)
        {
            return (data.NewData.IsRunning && !data.OldData.IsRunning) || (data.NewData.IsLaunching && !data.OldData.IsLaunching) ||
                   (!data.NewData.IsRunning && data.OldData.IsRunning) || (!data.NewData.IsLaunching && data.OldData.IsLaunching);
        }



        /// <summary>
        /// Looks for a <see cref="Guid"/> inside []-brackets inside a string.
        /// </summary>
        /// <param name="description">string containing the <see cref="Guid"/>.</param>
        /// <param name="gameId">The extracted <see cref="Guid"/>.</param>
        /// <returns>Whether a valid <see cref="Guid"/> could be parsed.</returns>
        private static bool ExtractIdFromLnkDescription(string description, out Guid gameId)
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
            }
            else
            {
                gameId = default;
                return false;
            }
        }

        public void UpdateShortcuts(IEnumerable<Game> games, ShortcutSyncSettings settings, bool forceUpdate = false)
        {
            CreateShortcuts(from game in games where ShouldKeepShortcut(game, settings) select game, settings);
            RemoveShortcuts(from game in games where !ShouldKeepShortcut(game, settings) select game, settings);
            if (forceUpdate) foreach (var game in games) if (existingShortcuts.ContainsKey(game.Id)) existingShortcuts[game.Id].Update(true);
        }

        public void CreateShortcuts(IEnumerable<Game> games, ShortcutSyncSettings settings)
        {
            foreach (var game in games)
            {
                if (existingShortcuts.ContainsKey(game.Id))
                {
                    if (!existingShortcuts[game.Id].Exists)
                    {
                        existingShortcuts[game.Id].Remove();
                        existingShortcuts.Remove(game.Id);
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
                        var copy = PlayniteApi.Database.Games.Get(copyId);
                        if (copy != null && copyId != game.Id)
                        {
                            if (existingShortcuts.ContainsKey(copyId))
                            {
                                existingShortcuts[copyId].Name = Path.GetFileNameWithoutExtension(GetShortcutPath(copy, settings.ShortcutPath, true, settings.SeparateFolders));
                                hasDuplicates = true;
                            }
                        }
                    }
                    shortcutNameToGameId[game.Name.GetSafeFileName().ToLower()].AddMissing(game.Id);
                }
                else
                {
                    shortcutNameToGameId[game.Name.GetSafeFileName().ToLower()] = new List<Guid>() { game.Id };
                }
                if (existingShortcuts.TryGetValue(game.Id, out Shortcut<Game> existing))
                {
                    existing.Name = Path.GetFileNameWithoutExtension(GetShortcutPath(game, settings.ShortcutPath, hasDuplicates, settings.SeparateFolders));
                }
                else
                {
                    if (settings.EnabledPlayActions.ContainsKey(GetSourceName(game)) && settings.EnabledPlayActions[GetSourceName(game)] && game.PlayAction != null)
                    {
                        string workingDirectory = PlayniteApi.ExpandGameVariables(game, game.PlayAction.WorkingDir);
                        string targetPath = PlayniteApi.ExpandGameVariables(game, game.PlayAction.Path);
                        string arguments = game.PlayAction.Arguments;
                        if (game.PlayAction.Type == GameActionType.Emulator)
                        {
                            var profile = PlayniteApi.Database.Emulators.Get(game.PlayAction.EmulatorId).Profiles.First(p => p.Id == game.PlayAction.EmulatorProfileId);
                            targetPath = profile.Executable;
                            workingDirectory = profile.WorkingDirectory;
                            arguments = "\"" + PlayniteApi.ExpandGameVariables(game, profile.Arguments) + "\"";
                        } 
                        existingShortcuts.Add(game.Id,
                            new TiledShortcutsPlayAction
                            (
                                targetGame: game,
                                shortcutPath: GetShortcutPath(game, settings.ShortcutPath, hasDuplicates, settings.SeparateFolders),
                                launchScriptFolder: GetLauncherScriptPath(),
                                tileIconFolder: GetLauncherScriptIconsPath(),
                                workingDirectory,
                                targetPath,
                                arguments
                            )
                        );
                    }
                    else
                    {
                        existingShortcuts.Add(game.Id,
                            new TiledShortcut
                            (
                                targetGame: game,
                                shortcutPath: GetShortcutPath(game, settings.ShortcutPath, hasDuplicates, settings.SeparateFolders),
                                launchScriptFolder: GetLauncherScriptPath(),
                                tileIconFolder: GetLauncherScriptIconsPath()
                            )
                        );
                    }
                }
            }
            // Stopwatch stopwatch = new Stopwatch();
            // stopwatch.Start();
            Parallel.ForEach(games, game => { if (existingShortcuts.ContainsKey(game.Id)) existingShortcuts[game.Id].CreateOrUpdate(); });
            // stopwatch.Stop();
            // PlayniteApi.Dialogs.ShowMessage($"Created {Shortcuts.Count} shortcuts in {stopwatch.ElapsedMilliseconds / 1000f} seconds.");
            //foreach (var game in games)
            {
                //  Shortcuts[game.Id].CreateOrUpdate();
            }
        }

        public void RemoveShortcuts(IEnumerable<Game> games, ShortcutSyncSettings settings)
        {
            foreach (var game in games)
            {
                if (game != null && existingShortcuts.ContainsKey(game.Id))
                {
                    existingShortcuts[game.Id].Remove();
                    existingShortcuts.Remove(game.Id);
                }
            }
            foreach (var game in games)
            {
                if (game == null) continue;
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
                            if (copy != null && existingShortcuts.ContainsKey(id))
                                existingShortcuts[id].Name = Path.GetFileNameWithoutExtension(GetShortcutPath(copy, settings.ShortcutPath, false, settings.SeparateFolders));
                        }
                }

            }
        }

        public void RemoveFromShortcutDicts(Guid gameId)
        {
            if (existingShortcuts.TryGetValue(gameId, out var shortcut))
            {
                shortcut.Remove();
                existingShortcuts.Remove(gameId);
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

        public void UpdateShortcutDicts(string folderPath, ShortcutSyncSettings settings, string shortcutName = "")
        {
            existingShortcuts.Clear();
            shortcutNameToGameId.Clear();

            if (!Directory.Exists(folderPath)) return;

            var files = Directory.GetFiles(folderPath, shortcutName + "*.lnk", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                var lnk = OpenLnk(file);
                if (lnk != null && ExtractIdFromLnkDescription(lnk.StringData.NameString, out Guid gameId))
                {
                    var game = PlayniteApi.Database.Games.Get(gameId);
                    if (game != null)
                    {
                        if (existingShortcuts.TryGetValue(game.Id, out var shortcut))
                        {
                            shortcut.Remove();
                        }
                        if (settings.EnabledPlayActions.ContainsKey(GetSourceName(game)) && settings.EnabledPlayActions[GetSourceName(game)] && game.PlayAction != null)
                        {
                            string workingDirectory = PlayniteApi.ExpandGameVariables(game, game.PlayAction.WorkingDir);
                            string targetPath = PlayniteApi.ExpandGameVariables(game, game.PlayAction.Path);
                            existingShortcuts.Add(game.Id,
                                new TiledShortcutsPlayAction
                                (
                                    targetGame: game,
                                    shortcutPath: file,
                                    launchScriptFolder: GetLauncherScriptPath(),
                                    tileIconFolder: GetLauncherScriptIconsPath(),
                                    workingDirectory,
                                    targetPath,
                                    game.PlayAction.Arguments
                                )
                            );
                        }
                        else
                        {
                            existingShortcuts[game.Id] = new TiledShortcut(game, file, GetLauncherScriptPath(), GetLauncherScriptIconsPath());
                        }
                        string safeGameName = game.Name.GetSafeFileName().ToLower();
                        if (!shortcutNameToGameId.ContainsKey(safeGameName))
                        {
                            shortcutNameToGameId.Add(safeGameName, new List<Guid>());
                        }
                        shortcutNameToGameId[safeGameName].AddMissing(gameId);
                    }
                }
            }
        }

        /// <summary>
        /// Opens/Creates and returns a <see cref="ShellLink.Shortcut"/> at a 
        /// given path.
        /// </summary>
        /// <param name="shortcutPath">Full path to the shortcut.</param>
        /// <returns>The shortcut object.</returns>
        public ShellLink.Shortcut OpenLnk(string shortcutPath)
        {
            return ShellLink.Shortcut.ReadFromFile(shortcutPath);
        }

        /// <summary>
        /// Checks whether a folder can be written to.
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns>Whether the folder can be written to.</returns>
        private static bool FolderIsAccessible(string folderPath)
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
        public static string GetSourceName(Game game)
        {
            if (game.Source == null)
            {
                return Constants.UNDEFINEDSOURCE;
            }
            else
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
            }
            else
            {
                logger.Debug($"Succesfully compiled launcher for {game.Name}");
                return true;
            }
        }


        private string GetLauncherScriptPath()
        {
            return Path.Combine(GetPluginUserDataPath(), Constants.LAUNCHSCRIPTFOLDERNAME);
        }

        private string GetLauncherScriptIconsPath()
        {
            return Path.Combine(GetPluginUserDataPath(), Constants.LAUNCHSCRIPTFOLDERNAME, Constants.ICONFOLDERNAME);
        }

        private bool CreateFolderStructure(string path)
        {
            try
            {
                Directory.CreateDirectory(path);
                var launchScriptsFolder = Directory.CreateDirectory(GetLauncherScriptPath());
                Directory.CreateDirectory(GetLauncherScriptIconsPath());
                launchScriptsFolder.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Could not create folder structure at \"{path}\".");
                return false;
            }
        }

        private static bool MergeMoveDirectory(string source, string target, bool overwrite = false)
        {
            if (Directory.Exists(target))
            {
                foreach (var dir in Directory.GetDirectories(source))
                {
                    MergeMoveDirectory(dir, Path.Combine(target, Path.GetFileName(dir)));
                }
                foreach (var file in Directory.GetFiles(source))
                {
                    var newFile = Path.Combine(target, Path.GetFileName(file));
                    if (!System.IO.File.Exists(newFile))
                    {
                        System.IO.File.Move(file, newFile);
                    }
                    else if (overwrite)
                    {
                        System.IO.File.Delete(newFile);
                        System.IO.File.Move(file, newFile);
                    }
                }
                Directory.Delete(source, true);
            }
            else
            {
                Directory.Move(source, target);
            }
            return true;
        }

        private string GetIconPath(Game game)
        {
            return Path.Combine(PlayniteApi.Database.GetFileStoragePath(game.Id), Path.GetFileName(game.Icon));
        }

        private bool HasExistingShortcutDuplicates(Game game)
        {
            if (shortcutNameToGameId.TryGetValue(game.Name.GetSafeFileName().ToLower(), out var copys))
            {
                foreach (var copy in copys)
                    if (game.Id != copy)
                        return true;
            }
            return false;
        }
    }
}