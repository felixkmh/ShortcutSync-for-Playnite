# ShortcutSync-for-Playnite

Plugin for [Playnite](https://github.com/JosefNemec/Playnite) that synchronizes shortcuts for all games recognized by Playnite so that they can be found by other launchers, e.g. Windows Search.

Can be set up to update on startup, otherwise it reacts to changes in the database, or it can be triggered manually. Shortcuts are only created for sources selected in the plugin settings. By default shortcuts are placed into a user's Start Menu folder. The Shortcuts point to a .vbs file that launches playnite with the corresponding game's id to run it or it uses a games PlayAction. Next to the .vbs file, an .xml file is created that allows a custom tile to be created when the shortcut is pinned to the Windows 10 start menu. 

Uses [IconLib](https://www.codeproject.com/Articles/16178/IconLib-Icons-Unfolded-MultiIcon-and-Windows-Vista) to extract the logo from .ico files.

[![Crowdin](https://badges.crowdin.net/shortcutsync-for-playnite/localized.svg)](https://crowdin.com/project/shortcutsync-for-playnite)

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/C1C6CH5IN)
