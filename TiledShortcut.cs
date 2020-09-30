using Playnite.SDK.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Drawing.IconLib;

namespace ShortcutSync
{
    public class TiledShortcut : Shortcut<Game>
    {
        public static string FileDatabasePath { get; set; } = null;

        public string LaunchScriptFolder    { get; protected set; } = null;
        public string TileIconFolder        { get; protected set; } = null;

        public TiledShortcut(Game targetGame, string shortcutPath, string launchScriptFolder, string tileIconFolder)
        {
            TargetObject = targetGame;
            LaunchScriptFolder = launchScriptFolder;
            TileIconFolder = tileIconFolder;
            ShortcutPath = shortcutPath;
        }

        public override bool Exists { 
            get
            {
                bool shortcutExists = !ShortcutPath.IsNullOrEmpty() && File.Exists(ShortcutPath);
                bool launchScriptExists = !GetLauncherPath().IsNullOrEmpty() && File.Exists(GetLauncherPath());
                bool shortcutIconExists = GetGameIconPath().IsNullOrEmpty() || File.Exists(GetShortcutIconPath());
                bool tileIconExists = GetGameIconPath().IsNullOrEmpty() || File.Exists(GetTileIconPath());
                return shortcutExists && launchScriptExists && shortcutIconExists && tileIconExists;
            }
        }

        public override DateTime LastUpdated {
            get 
            {
                if (File.Exists(ShortcutPath))
                {
                    return File.GetLastWriteTime(ShortcutPath);
                }
                return DateTime.MinValue;
            }
            protected set 
            { 
                throw new FieldAccessException("Not allowed to set LastUpdated property."); 
            } 
        }

        public override string Name { 
            get {
                return Path.GetFileNameWithoutExtension(ShortcutPath);            
            }
            set {
                if (!value.IsNullOrEmpty())
                    RenameShortcut(value);
            }  
        }

        public override bool IsOutdated
        {
            get
            {
                if (TargetObject.Modified is DateTime modified)
                {
                    return (modified > LastUpdated);
                }
                return false;
            }
        }

        protected void RenameShortcut(string name)
        {
            string newName = name.GetSafeFileName();
            if (!newName.IsNullOrEmpty())
            {
                string newShortcutPath = Path.Combine(Path.GetDirectoryName(ShortcutPath), newName + ".lnk");
                if (Exists && ShortcutPath != newShortcutPath)
                {
                        File.Move(ShortcutPath, newShortcutPath);
                }
                ShortcutPath = newShortcutPath;
            }
        }

        public override bool Create()
        {
            if (!Exists)
            {
                CreateFolderStructure();
                CreateTileImage();
                CreateVbsLauncher();
                CreateVisualElementsManifest();
                CreateLnk();
                return true;
            }
            else
            {
                return false;
            }
        }

        protected void CreateFolderStructure()
        {
            if (!Directory.Exists(Path.GetDirectoryName(ShortcutPath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(ShortcutPath));
            }
            if (!Directory.Exists(LaunchScriptFolder))
            {
                var scripts = Directory.CreateDirectory(LaunchScriptFolder);
                scripts.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
            if (!Directory.Exists(TileIconFolder))
            {
                Directory.CreateDirectory(TileIconFolder);
            }
        }

        protected void CreateLnk()
        {
            var lnk = OpenLnk(ShortcutPath);
            if (CreateShortcutIcon())
            {
                lnk.IconLocation = GetShortcutIconPath();
            }
            // TODO: Get right icon path
            lnk.TargetPath = GetLauncherPath();
            lnk.Description = "Launch " + TargetObject.Name + " on " + GetSourceName(TargetObject) + " via Playnite." + $" [{TargetObject.Id}]";
            lnk.WorkingDirectory = "";
            lnk.Save();
        }

        public override bool Remove()
        {
            SafeDelete(ShortcutPath);
            SafeDelete(GetLauncherPath());
            SafeDelete(GetManifestPath());
            SafeDelete(GetTileIconPath());
            return true;
        }

        public override bool Update(bool forceUpdate = false)
        {
            if (Exists)
            {
                if (IsOutdated || forceUpdate)
                {
                    // RenameShortcut(TargetObject.Name.GetSafeFileName());
                    CreateTileImage();
                    CreateLnk();
                    return true;
                }
            }
            return false;
        }

        public override bool Update(Game targetObject)
        {
            TargetObject = targetObject;
            return Update();
        }

        public override bool IsValid
        {
            get
            {
                bool gameIsValid = TargetObject != null;
                bool shortcutPathIsValid = Path.GetExtension(ShortcutPath).ToLower() == ".lnk";
                bool launchScriptPathIsValid = Path.GetExtension(GetLauncherPath()).ToLower() == ".vbs";
                bool tileIconIsValid = GetTileIconPath().IsNullOrEmpty() || Path.GetExtension(GetTileIconPath()).ToLower() == ".png";
                bool ShortcutIconIsValid = GetShortcutIconPath().IsNullOrEmpty() || Path.GetExtension(GetShortcutIconPath()).ToLower() == ".ico";
                return gameIsValid
                    && shortcutPathIsValid
                    && launchScriptPathIsValid
                    && tileIconIsValid
                    && ShortcutIconIsValid;
            }
        }
        
        protected string GetManifestName()
        {
            return $"{TargetObject.Id}.visualelementsmanifest.xml";
        }

        protected string GetManifestPath()
        {
            return Path.Combine(LaunchScriptFolder, GetManifestName());
        }

        protected string GetTileIconName()
        {
            return $"{TargetObject.Id}.png";
        }

        protected string GetTileIconPath()
        {
            return Path.Combine(TileIconFolder, GetTileIconName());
        }

        protected string GetLauncherName()
        {
            return $"{TargetObject.Id}.vbs";
        }

        protected string GetLauncherPath()
        {
            return Path.Combine(LaunchScriptFolder, GetLauncherName());
        }

        protected bool SafeDelete(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Opens/Creates and returns a <see cref="IWshShortcut"/> at a 
        /// given path.
        /// </summary>
        /// <param name="shortcutPath">Full path to the shortcut.</param>
        /// <returns>The shortcut object.</returns>
        public IWshRuntimeLibrary.IWshShortcut OpenLnk(string shortcutPath)
        {
            if (true)
            {
                try
                {
                    var shell = new IWshRuntimeLibrary.WshShell();
                    var shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcutPath);
                    return shortcut;
                }
                catch (Exception){}
            }
            return null;
        }


        /// <summary>
        /// Safely gets a name of the source for a game.
        /// </summary>
        /// <param name="game"></param>
        /// <returns>Source name or "Undefined" if no source exists.</returns>
        protected string GetSourceName(Game game)
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

        protected virtual string CreateVbsLauncher()
        {
            string fullPath = GetLauncherPath();
            string script =
                "Dim prefix, id\n" +

                "prefix = \"playnite://playnite/start/\"\n" +
                $"id = \"{TargetObject.Id}\"\n" +

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

        protected string CreateVisualElementsManifest()
        {
            string fullPath = GetManifestPath();
            string foregroundTextStyle = "light";
            string backgroundColorCode = "#000000";
            (var bgColor, var brightness) = CreateTileImage();
            foregroundTextStyle = brightness > 0.4f ? "dark" : "light";
            backgroundColorCode = bgColor.ToHexCode();
            string script =
                "<Application xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n" +
                "<VisualElements\n" +
                $"BackgroundColor = \"{backgroundColorCode}\"\n" +
                "ShowNameOnSquare150x150Logo = \"on\"\n" +
                $"ForegroundText = \"{foregroundTextStyle}\"\n" +
                $"Square150x150Logo = \"{Constants.ICONFOLDERNAME}\\{TargetObject.Id}.png\"\n" +
                $"Square70x70Logo = \"{Constants.ICONFOLDERNAME}\\{TargetObject.Id}.png\"/>\n" +
                "</Application>";
            try
            {
                using (var scriptFile = System.IO.File.CreateText(fullPath))
                {
                    scriptFile.Write(script);
                }
                return fullPath;
            }
            catch (Exception) { }

            return string.Empty;
        }

        
        protected (Color bgColor, float brightness) CreateTileImage()
        {
            string iconPath = GetGameIconPath();
            Color bgColor = Color.DarkGray;
            float brightness = 0.5f;
            if (!Path.GetFileName(TargetObject.Icon).IsNullOrEmpty()
                && System.IO.File.Exists(iconPath))
            {
                Bitmap bitmap = null;
                if (Path.GetExtension(TargetObject.Icon).ToLower() == ".ico")
                {
                    bitmap = ExtractBitmapFromIcon(iconPath, 150, 150);
                }
                else
                {
                    bitmap = new Bitmap(iconPath);
                }
                if (bitmap != null && bitmap.Width > 0 && bitmap.Height > 0)
                {
                    brightness = GetAverageBrightness(bitmap);
                    
                    bgColor = GetDominantColor(bitmap, brightness);

                    // resize
                    int newWidth = 150;
                    int newHeight = 150;
                    if (bitmap.Width >= bitmap.Height)
                    {
                        float scale = (float)newHeight / bitmap.Height;
                        newWidth = (int)Math.Round(bitmap.Width * scale);
                    }
                    else
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
                    var tileIconPath = Path.Combine(TileIconFolder, $"{TargetObject.Id}.png");
                    resized.Save(tileIconPath, ImageFormat.Png);
                    brightness = GetLowerThirdBrightness(resized, bgColor);
                    resized.Dispose();
                } 
                bitmap.Dispose();
            } 
            return (bgColor, brightness);
        }

        protected static Bitmap ExtractBitmapFromIcon(string iconPath, int desiredWidth = 150, int desiredHeight = 150)
        {
            Bitmap bitmap = null;
            using (var stream = System.IO.File.OpenRead(iconPath))
            {
                var i = new System.Drawing.IconLib.MultiIcon();
                i.Load(stream);
                int maxIndex = 0;
                int maxWidth = 0;
                int index = 0;
                foreach (var imag in i[0])
                {
                    if (imag.Size.Width >= desiredWidth && imag.Size.Height >= desiredHeight)
                    {
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
            }
            return bitmap;
        }

        protected float GetAverageBrightness(Bitmap bitmap)
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

        protected Color GetDominantColor(Bitmap bitmap, float brightness = 0.5f)
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
                    red: Math.Max(0, (int)Math.Min(255, r)),
                    green: Math.Max(0, (int)Math.Min(255, g)),
                    blue: Math.Max(0, (int)Math.Min(255, b))
                );
            var brightnessFactor = 0;
            if (Math.Abs(brightness) - Math.Abs(color.GetBrightness()) < 0.25)
                brightnessFactor = brightness >= color.GetBrightness() ? -1 : 1;
            color = System.Drawing.Color.FromArgb(
                    alpha: 255,
                    red: Math.Max(0, (int)Math.Min(255, color.R + brightnessFactor * 90)),
                    green: Math.Max(0, (int)Math.Min(255, color.G + brightnessFactor * 90)),
                    blue: Math.Max(0, (int)Math.Min(255, color.B + brightnessFactor * 90))
                );
            return color;
        }

        protected float GetLowerThirdBrightness(Bitmap bitmap, Color bgColor)
        {
            float bgBrightness = bgColor.GetBrightness();
            float brightness = 0f;
            int almostBlack = 0;
            int almostWhite = 0;
            int dark = 0;
            int bright = 0;
            int numberOfSamples = 0;
            for (int y = (int)Math.Round(bitmap.Height * (2f/3f)); y < bitmap.Height; ++y)
                for (int x = (int)Math.Round(bitmap.Width * (2f / 3f)); x < bitmap.Width; ++x)
                {
                    var pixelColor = bitmap.GetPixel(x, y);
                    float alpha = pixelColor.A / 255f;
                    float r = bgColor.R * (1f - alpha) + pixelColor.R * alpha;
                    float g = bgColor.G * (1f - alpha) + pixelColor.G * alpha;
                    float b = bgColor.B * (1f - alpha) + pixelColor.B * alpha;
                    var sampleColor = Color.FromArgb(255, (int)Math.Round(r), (int)Math.Round(g), (int)Math.Round(b));
                    var sample = sampleColor.GetBrightness();
                    brightness += sample;
                    almostBlack += sample <= 0.1 ? 1 : 0;
                    almostWhite += sample >= 0.9 ? 1 : 0;
                    dark += sample < 0.35f ? 1 : 0;
                    bright += sample > 0.65f ? 1 : 0;
                    ++numberOfSamples;
                }
            brightness /= numberOfSamples;
            if (almostWhite > 10 && almostBlack > 10)
            {
                brightness += 1.5f * almostWhite / numberOfSamples;
                brightness -= 1.5f * almostBlack / numberOfSamples;
            } else if (almostBlack >= almostWhite)
            {
                brightness = 0;
            } else
            {
                brightness = 1;
            }
            return Math.Max(0, Math.Min(1, (float)bright / (dark + bright)));
        }

        protected string GetGameIconPath()
        {
            return TargetObject.Icon.IsNullOrEmpty() 
                ? string.Empty 
                : Path.Combine(FileDatabasePath, TargetObject.Id.ToString(), Path.GetFileName(TargetObject.Icon));
        }

        protected string GetShortcutIconPath()
        {
            return GetGameIconPath().IsNullOrEmpty()
                ? string.Empty
                : Path.ChangeExtension(GetGameIconPath(), ".ico");
        }

        protected bool LoadManifest(out XmlDocument manifest)
        {
            manifest = null;
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            try { doc.Load(GetLauncherPath()); }
            catch
            {
                return false;
            }
            manifest = doc;
            return true;
        }

        protected bool CreateShortcutIcon()
        {
            string gameIconPath = GetGameIconPath();
            // Skip if a suitable icon file already exists
            if (Path.GetExtension(gameIconPath).ToLower() == ".ico")
                return true;

            if (File.Exists(gameIconPath))
            {
                using (Bitmap bitmap = new Bitmap(gameIconPath))
                {
                    Size iconSize = new Size(256, 256);
                    using (Bitmap resized = new Bitmap(iconSize.Width, iconSize.Height))
                    {
                        using (Graphics g = Graphics.FromImage(resized))
                        {
                            g.DrawImage(bitmap, 0, 0, iconSize.Width, iconSize.Height);
                        }

                        MultiIcon mIcon = new MultiIcon();
                        SingleIcon sIcon = mIcon.Add("Original");
                        sIcon.CreateFrom(resized, IconOutputFormat.FromWinXP);
                        mIcon.SelectedIndex = 0;
                        mIcon.Save(GetShortcutIconPath(), MultiIconFormat.ICO);
                        return true;

                    }
                }
            }
            return false;
        }

        public override bool Move(string path)
        {
            throw new NotImplementedException();
        }
    }
}
