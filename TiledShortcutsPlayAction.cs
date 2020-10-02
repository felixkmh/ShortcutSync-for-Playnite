using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Playnite.SDK.Models;

namespace ShortcutSync
{
    public class TiledShortcutsPlayAction : TiledShortcut
    {
        public string WorkingDir { get; protected set; } = "";
        public string TargetPath { get; protected set; } = "";
        public string Arguments { get; protected set; } = "";

        public TiledShortcutsPlayAction(Game targetGame, string shortcutPath, string launchScriptFolder, string tileIconFolder)
            : base(targetGame, shortcutPath, launchScriptFolder, tileIconFolder)
        {}

        public TiledShortcutsPlayAction(Game targetGame, string shortcutPath, string launchScriptFolder, string tileIconFolder, string workinkDir, string targetPath, string arguments)
            : base(targetGame, shortcutPath, launchScriptFolder, tileIconFolder)
        {
            WorkingDir = workinkDir;
            TargetPath = targetPath;
            Arguments = arguments;
        }

        protected override string CreateVbsLauncher()
        {
            string fullPath = GetLauncherPath();
            string script = "";
            if (TargetObject.Source != null && TargetObject.Source.Name == "Xbox")
            {
                script =
                "Set WshShell = WScript.CreateObject(\"WScript.Shell\")\n" +
                $"WshShell.CurrentDirectory = \"Applications\"\n" +
                $"WshShell.Run \"{@"C:Windows\explorer.exe"}\" & \" \" & \"{Arguments}\" , 1";
            } else if (TargetObject.PlayAction.Type == GameActionType.URL)
            {
                script =
                "Set WshShell = WScript.CreateObject(\"WScript.Shell\")\n" +
                $"WshShell.Run \"{TargetObject.PlayAction.Path}\", 1";
            }
            else
            {
                script =
                "Set WshShell = WScript.CreateObject(\"WScript.Shell\")\n" +
                $"WshShell.CurrentDirectory = \"{WorkingDir}\"\n" +
                $"WshShell.Run \"{TargetObject}\" & \" \" & \"{Arguments}\" , 1";
            }

          
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
    }
}
