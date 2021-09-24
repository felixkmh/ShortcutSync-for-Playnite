using Playnite.SDK.Models;
using System;
using System.Drawing;
using System.Drawing.IconLib;
using System.Drawing.Imaging;
using System.IO;
using System.Xml;
using System.Numerics;
using KamilSzymborski.VisualElementsManifest;
using KamilSzymborski.VisualElementsManifest.Extensions;
using KamilSzymborski.VisualElementsManifest.Tools;
using System.Text;
using System.Linq;

namespace ShortcutSync
{
    public class TiledShortcut : Shortcut<Game>
    {
        public static string FileDatabasePath => Path.Combine(ShortcutSync.Instance.PlayniteApi.Database.DatabasePath, "files");
        public static string DefaultIconPath => Path.Combine(ShortcutSync.Instance.PlayniteApi.Paths.ApplicationPath, "Themes", "Desktop", "Default", "Images", "applogo.png");
        public static bool FadeTileEdge { get; set; } = false;

        public string LaunchScriptFolder { get; protected set; } = null;
        public string TileIconFolder { get; protected set; } = null;

        public TiledShortcut(Game targetGame, string shortcutPath, string launchScriptFolder, string tileIconFolder)
        {
            TargetObject = targetGame;
            LaunchScriptFolder = launchScriptFolder;
            TileIconFolder = tileIconFolder;
            ShortcutPath = shortcutPath;
        }

        public override bool Exists
        {
            get
            {
                bool shortcutExists = !ShortcutPath.IsNullOrEmpty() && File.Exists(ShortcutPath);
                bool launchScriptExists = !GetLauncherPath().IsNullOrEmpty() && File.Exists(GetLauncherPath());
                bool shortcutIconExists = GetGameIconPath().IsNullOrEmpty() || File.Exists(GetShortcutIconPath());
                bool tileIconExists = GetGameIconPath().IsNullOrEmpty() || File.Exists(GetTileIconPath());
                return shortcutExists && launchScriptExists && shortcutIconExists && tileIconExists;
            }
        }

        public override DateTime LastUpdated
        {
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

        public override string Name
        {
            get
            {
                return Path.GetFileNameWithoutExtension(ShortcutPath);
            }
            set
            {
                if (!value.IsNullOrEmpty())
                    if (Path.GetFileNameWithoutExtension(ShortcutPath) != value)
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
                CreateShortcutIcon();
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
            var link = ShellLink.Shortcut.CreateShortcut(GetLauncherPath());
            link.IconIndex = 0;
            link.StringData = new ShellLink.Structures.StringData(true)
            {
                IconLocation = GetShortcutIconPath(),
                NameString = "Launch " + TargetObject.Name + " on " + GetSourceName(TargetObject) + " via Playnite." + $" [{TargetObject.Id}]"
            };
            link.StringData.NameString = "Launch " + TargetObject.Name + " on " + GetSourceName(TargetObject) + " via Playnite." + $" [{TargetObject.Id}]";
            link.WriteToFile(ShortcutPath);
        }

        protected ShellLink.Shortcut OpenLnk()
        {
            return ShellLink.Shortcut.ReadFromFile(ShortcutPath);
        }

        public override bool Remove()
        {

            bool shortcutDeleted = SafeDelete(ShortcutPath);
            bool launcherScriptDeleted = SafeDelete(GetLauncherPath());
            bool manifestDeleted = SafeDelete(GetManifestPath());
            bool iconDeleted = SafeDelete(GetTileIconPath());
            bool smallIconDeleted = SafeDelete(GetTileSmallIconPath());

            if (iconDeleted && smallIconDeleted && Directory.GetFileSystemEntries(TileIconFolder).Length == 0)
                Directory.Delete(TileIconFolder);

            if (launcherScriptDeleted && manifestDeleted && Directory.GetFileSystemEntries(LaunchScriptFolder).Length == 0)
                Directory.Delete(LaunchScriptFolder);

            if (shortcutDeleted && Directory.GetFileSystemEntries(Path.GetDirectoryName(ShortcutPath)).Length == 0)
                Directory.Delete(Path.GetDirectoryName(ShortcutPath));

            return true;
        }

        public override bool Update(bool forceUpdate = false)
        {
            if (Exists)
            {
                if (IsOutdated || forceUpdate)
                {
                    CreateShortcutIcon();
                    CreateTileImage();
                    CreateVbsLauncher();
                    CreateVisualElementsManifest();
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

        protected string GetTileSmallIconName()
        {
            return $"{TargetObject.Id}_70.png";
        }

        protected string GetTileSmallIconPath()
        {
            return Path.Combine(TileIconFolder, GetTileSmallIconName());
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
                "WshShell.Run prefix & id, 1\n"+
                "Set WshShell=Nothing"; ;
            try
            {
                using (var scriptFile = new StreamWriter(fullPath, false))
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
            if (File.Exists(fullPath))
            {
                var xml = File.ReadAllText(fullPath);
                if (ManifestService.Validate(xml))
                {
                    var manifest = ManifestService.Parse(xml);
                    xml = ManifestService.Create(manifest);
                    if (!File.Exists(Path.Combine(LaunchScriptFolder, manifest.GetSetSquare150x150Logo())))
                    {
                        manifest.SetSquare150x150LogoOn(Path.Combine(Constants.ICONFOLDERNAME, TargetObject.Id + ".png"));
                    }
                    if (!File.Exists(Path.Combine(LaunchScriptFolder, manifest.GetSetSquare70x70Logo())))
                    {
                        manifest.SetSquare70x70LogoOn(Path.Combine(Constants.ICONFOLDERNAME, TargetObject.Id + "_70.png"));
                    }
                    var newXml = ManifestService.Create(manifest);
                    if (newXml != xml)
                    {
                        File.WriteAllText(fullPath, newXml);
                    }
                    return fullPath;
                }
            } 

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
                $"Square70x70Logo = \"{Constants.ICONFOLDERNAME}\\{TargetObject.Id}_70.png\"/>\n" +
                "</Application>";
            try
            {
                using (var scriptFile = new StreamWriter(fullPath, false))
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
            if (System.IO.File.Exists(iconPath))
            {
                // Check if file can be accessed
                try
                {
                    using (FileStream stream = File.Open(iconPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        stream.Close();
                    }
                }
                catch (IOException ex)
                {
                    ShortcutSync.logger?.Error(ex, $"Could not open icon at \"{iconPath}\" to create tile image for game {TargetObject?.Name??""}.");
                    return (bgColor, brightness);
                }
                Bitmap bitmap = null;
                if (Path.GetExtension(iconPath).ToLower() == ".ico")
                {
                    bitmap = ExtractBitmapFromIcon(iconPath, 150, 150);
                }
                if (bitmap == null)
                {
                    bitmap = new Bitmap(iconPath);
                }
                if (bitmap != null && bitmap.Width > 0 && bitmap.Height > 0)
                {
                    brightness = GetAverageBrightness(bitmap);

                    bgColor = GetDominantColor(bitmap, brightness);

                    // resize 70x70
                    int newWidth = 70;
                    int newHeight = 70;
                    if (bitmap.Width <= bitmap.Height)
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
                        graphics.RenderingOrigin = new Point(34, 34);
                        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
                        if (bitmap.Width <= 16 && bitmap.Height <= 16)
                        {
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                            graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                            graphics.DrawImage(bitmap, new RectangleF(0, 0, 70, 70), new RectangleF(0, 0, bitmap.Width, bitmap.Height), GraphicsUnit.Pixel);
                        }
                        else
                        {
                            graphics.DrawImage(bitmap, new Rectangle(0, 0, newWidth, newHeight));
                        }
                    }
                    resized.Save(GetTileSmallIconPath(), ImageFormat.Png);
                    resized.Dispose();

                    // resize 150x150
                    newWidth = 150;
                    newHeight = 150;
                    if (bitmap.Width <= bitmap.Height)
                    {
                        float scale = (float)newHeight / bitmap.Height;
                        newWidth = (int)Math.Round(bitmap.Width * scale);
                    }
                    else
                    {
                        float scale = (float)newWidth / bitmap.Width;
                        newHeight = (int)Math.Round(bitmap.Height * scale);
                    }
                    resized = new Bitmap(newWidth, newHeight);
                    using (Graphics graphics = Graphics.FromImage(resized))
                    {
                        graphics.RenderingOrigin = new Point(74, 74);
                        if (bitmap.Width <= 64 && bitmap.Height <= 64)
                        {
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                            graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                            graphics.DrawImage(bitmap, new RectangleF(0, 0, 150, 150), new RectangleF(0, 0, bitmap.Width, bitmap.Height), GraphicsUnit.Pixel);
                        } else
                        {
                            graphics.DrawImage(bitmap, new Rectangle(0, 0, newWidth, newHeight));
                        }
                    }
                    var tileIconPath = GetTileIconPath();
                    var showTextOnTile = true;
                    if (File.Exists(GetManifestPath()))
                    {
                        var manifest = ManifestService.Parse(File.ReadAllText(GetManifestPath()));
                        showTextOnTile = manifest.IsShowNameOnSquare150x150LogoSetOnOn();
                    }
                    if (FadeTileEdge && showTextOnTile)
                    {
                        float percentage = 0.28f;
                        if (TargetObject.Name.Length >= 12 && TargetObject.Name.Split(null).Length > 1)
                        {
                            percentage = 0.42f;
                        }
                        BlendEdge(resized, Edge.Top, percentage, 2.8f, 0.5f);
                    }
                    resized.Save(tileIconPath, ImageFormat.Png);
                    brightness = GetLowerThirdBrightness(resized, bgColor);
                    resized.Dispose();
                }
                bitmap.Dispose();
            }
            return (bgColor, brightness);
        }

        protected enum Edge
        {
            Bottom,
            Left,
            Top,
            Right
        }

        protected static void BlendEdge(Bitmap bitmap, Edge edge, float percentage, float exponent = 1f, float minAlpha = 0f)
        {
            Vector2 start = new Vector2();
            Vector2 dir = new Vector2();
            float length = 0f;
            switch (edge)
            {
                case Edge.Bottom:
                    dir.Y = 1;
                    length = percentage * bitmap.Height;
                    break;
                case Edge.Left:
                    dir.X = 1;
                    length = percentage * bitmap.Width;
                    break;
                case Edge.Top:
                    start.Y = bitmap.Height;
                    dir.Y = -1;
                    length = percentage * bitmap.Height;
                    break;
                case Edge.Right:
                    start.X = bitmap.Width;
                    dir.X = -1;
                    length = percentage * bitmap.Width;
                    break;
                default:
                    break;
            }
            for(int y = 0; y < bitmap.Height; ++y)
            {
                for (int x = 0; x < bitmap.Width; ++x)
                {
                    Vector2 toPoint = new Vector2(x, y) - start;
                    float alpha = Math.Max(0f, Math.Min(1f, Vector2.Dot(dir, toPoint) / length));
                    alpha = (float)Math.Pow(alpha, exponent);
                    Color oldColor = bitmap.GetPixel(x, y);
                    // Color newColor = oldColor.BlendAlpha(blendColor, alpha);
                    Color newColor = Color.FromArgb(Math.Min((int)Math.Round(alpha.Map(0, 1, minAlpha, 1) * 255), oldColor.A), oldColor.R, oldColor.G, oldColor.B);
                    bitmap.SetPixel(x, y, newColor);
                }
            }
        }

        protected static Bitmap ExtractBitmapFromIcon(string iconPath, int desiredWidth = 150, int desiredHeight = 150)
        {
            Bitmap bitmap = null;
            var i = new System.Drawing.IconLib.MultiIcon();
            var si = i.Add("icon");
            using (var stream = new FileStream(iconPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                si.Load(stream);
            }
            int maxWidth = 0;
            int maxIdx = -1;
            int idx = 0;
            foreach (var imag in si)
            {
                // imag.IconImageFormat = IconImageFormat.PNG;
                if (imag.Size.Width >= desiredWidth && imag.Size.Height >= desiredHeight && imag.IconImageFormat == IconImageFormat.PNG)
                {
                    return imag.Transparent;
                }
                if (imag.Size.Width > maxWidth && imag.IconImageFormat == IconImageFormat.PNG)
                {
                    maxIdx = idx;
                    maxWidth = imag.Size.Width;
                }
                ++idx;
            }
            if (maxIdx > -1)
            {
                bitmap = si[maxIdx].Transparent;
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
            for (int y = (int)Math.Round(bitmap.Height * (2f / 3f)); y < bitmap.Height; ++y)
                for (int x = 0; x < bitmap.Width; ++x)
                {
                    var pixelColor = bitmap.GetPixel(x, y);
                    float alpha = pixelColor.A / 255f;
                    if (alpha > 0.75)
                    {
                        float r = bgColor.R * (1f - alpha) + pixelColor.R * alpha;
                        float g = bgColor.G * (1f - alpha) + pixelColor.G * alpha;
                        float b = bgColor.B * (1f - alpha) + pixelColor.B * alpha;
                        var sampleColor = Color.FromArgb(255, (int)Math.Round(r), (int)Math.Round(g), (int)Math.Round(b));
                        var sample = pixelColor.GetBrightness();
                        brightness += sample;
                        almostBlack += sample <= 0.1 ? 1 : 0;
                        almostWhite += sample >= 0.9 ? 1 : 0;
                        dark += sample <= 0.5f ? 1 : 0;
                        bright += sample > 0.5f ? 1 : 0;
                        ++numberOfSamples;
                    }
                }
            brightness /= numberOfSamples;
            if (almostWhite > 10 && almostBlack > 10)
            {
                brightness += 1.5f * almostWhite / numberOfSamples;
                brightness -= 1.5f * almostBlack / numberOfSamples;
            }
            else if (almostBlack >= almostWhite)
            {
                brightness = 0;
            }
            else
            {
                brightness = 1;
            }
            return dark > bright ? 0 : 1;
        }

        protected string GetGameIconPath()
        {
            var databasePath = TargetObject.Icon;
            if (!string.IsNullOrEmpty(databasePath))
            {
                var path = ShortcutSync.Instance.PlayniteApi.Database.GetFullFilePath(TargetObject.Icon);
                if (databasePath.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    // URL: TODO

                } else if (File.Exists(path))
                {
                    // File
                    return path;
                }
            }
            return Path.ChangeExtension(DefaultIconPath, ".png");
        }

        protected string GetShortcutIconPath()
        {
            return Path.ChangeExtension(GetGameIconPath(), ".ico");
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
            // Skip if a suitable icon file already exists
            var originalPath = GetShortcutIconPath();
            if (File.Exists(originalPath))
            {
                var bytes = new byte[4];
                using (var file = File.OpenRead(originalPath))
                {
                    file.Read(bytes, 0, 4);
                }
                if (bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 1 && bytes[3] == 0)
                {
                    return true;
                }
                using (var img = new Bitmap(originalPath))
                {
                    var extension = ImageCodecInfo.GetImageDecoders().FirstOrDefault(dec => dec.FormatID == img.RawFormat.Guid)?.FilenameExtension.Replace("*",string.Empty).ToLower().Split(';');
                    if (extension != null)
                    {
                        if (!extension.Contains(Path.GetExtension(originalPath).ToLower()))
                        {
                            FixWrongIconExtension(originalPath, img);
                        }
                    }
                }
            }

            string gameIconPath = GetGameIconPath();

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
                        sIcon.CreateFrom(resized, IconOutputFormat.FromWin95);
                        mIcon.SelectedIndex = 0;
                        mIcon.Save(GetShortcutIconPath(), MultiIconFormat.ICO);
#if DEBUG
                        // if (File.Exists(GetGameIconPath())) throw new Exception($"{GetShortcutIconPath()} was not created!");
#endif
                        return true;

                    }
                }
            }
            return false;
        }

        private void FixWrongIconExtension(string file, Image img)
        {
            if (TargetObject != null)
            {
                var extensions = ImageCodecInfo.GetImageDecoders().FirstOrDefault(enc => enc.FormatID == img.RawFormat.Guid)?.FilenameExtension;
                if (extensions is string)
                {
                    var ext = extensions.Split(';').FirstOrDefault()?.ToLower().Replace("*", string.Empty);
                    if (ext is string)
                    {
                        var tempPath = Path.GetTempFileName() + ext;
                        File.Copy(file, tempPath);
                        ShortcutSync.Instance.PlayniteApi.Database.RemoveFile(Path.GetFileNameWithoutExtension(file));
                        var id = ShortcutSync.Instance.PlayniteApi.Database.AddFile(tempPath, TargetObject.Id);
                        File.Delete(tempPath);
                        TargetObject.Icon = id;
                        ShortcutSync.Instance.PlayniteApi.Database.Games.Update(TargetObject);
                    }
                }
            }
        }

        public override bool Move(params string[] paths)
        {
            if (paths.Length != 3)
            {
                throw new ArgumentException("TiledShortcut.Move needs three paths: shortcutPath, tileIconPath and launchScriptPath.", nameof(paths));
            }

            if (!Exists) return false;

            string newShortcutPath = paths[0];
            string newTileIconPath = paths[1];
            string newLaunchScriptPath = paths[2];

            var newShortcutPathType = newShortcutPath.GetPathType();
            var newTileIconPathType = newTileIconPath.GetPathType();
            var newLaunchScriptPathType = newLaunchScriptPath.GetPathType();

            if (newShortcutPathType == Extensions.PathType.Invalid)
                throw new ArgumentException($"{newShortcutPath} is not a valid path.", nameof(newShortcutPath));
            if (newTileIconPathType != Extensions.PathType.Directory)
                throw new ArgumentException($"{newTileIconPath} is not a valid directory.", nameof(newTileIconPath));
            if (newLaunchScriptPathType != Extensions.PathType.Directory)
                throw new ArgumentException($"{newLaunchScriptPath} is not a valid directory.", nameof(newLaunchScriptPath));

            if (!TargetObject.Icon.IsNullOrEmpty())
                if (MoveFileToFileOrDirectory(
                    Path.Combine(TileIconFolder, TargetObject.Id + ".png"), 
                    Path.Combine(newTileIconPath, TargetObject.Id + ".png")))
                {
                    if (Directory.GetFileSystemEntries(TileIconFolder).Length == 0)
                        Directory.Delete(TileIconFolder);
                    TileIconFolder = newTileIconPath;
                }
                else
                {
                    return false;
                }

            bool moved = false;
            if (MoveFileToFileOrDirectory(
                Path.Combine(LaunchScriptFolder, TargetObject.Id + ".vbs"), 
                Path.Combine(newLaunchScriptPath, TargetObject.Id + ".vbs")))
            {
                if (Directory.GetFileSystemEntries(LaunchScriptFolder).Length == 0)
                    Directory.Delete(LaunchScriptFolder);
                moved = true;
            }
            else
            {
                return false;
            }

            if (MoveFileToFileOrDirectory(
                Path.Combine(LaunchScriptFolder, TargetObject.Id + ".visualelementsmanifest.xml"),
                Path.Combine(newLaunchScriptPath, TargetObject.Id + ".visualelementsmanifest.xml")
               ))
            {
                if (Directory.GetFileSystemEntries(LaunchScriptFolder).Length == 0)
                    Directory.Delete(LaunchScriptFolder);
                if (moved)
                {
                    LaunchScriptFolder = newLaunchScriptPath;
                    var lnk = OpenLnk();
                    if (lnk.StringData.RelativePath != GetLauncherPath())
                    {
                        lnk.StringData.RelativePath = GetLauncherPath();
                        lnk.ExtraData.EnvironmentVariableDataBlock.TargetAnsi = GetLauncherPath();
                        lnk.ExtraData.EnvironmentVariableDataBlock.TargetUnicode = GetLauncherPath();
                        lnk.WriteToFile(ShortcutPath);
                    }
                }
            }
            else
            {
                return false;
            }

            if (MoveFileToFileOrDirectory(ShortcutPath, newShortcutPath))
            {
                if (Directory.GetFileSystemEntries(Path.GetDirectoryName(ShortcutPath)).Length == 0)
                    Directory.Delete(Path.GetDirectoryName(ShortcutPath));
                if (newShortcutPathType == Extensions.PathType.File)
                {
                    ShortcutPath = newShortcutPath;
                }
                else
                {
                    ShortcutPath = Path.Combine(newShortcutPath, Path.GetFileName(ShortcutPath));
                }
            }
            else
            {
                return false;
            }


            return true;
        }

        protected bool MoveFileToFileOrDirectory(string source, string target, bool overwrite = false)
        {
            var sourceType = source.GetPathType();
            var targetType = source.GetPathType();

            string newPath = target;
            if (targetType == Extensions.PathType.Directory)
                newPath = Path.Combine(target, Path.GetFileName(source));

            if (source == newPath) return true;

            if (sourceType != Extensions.PathType.File) return false;
            if (!File.Exists(source)) return false;

            if (File.Exists(newPath))
            {
                if (overwrite)
                    File.Delete(target);
                else
                    return false;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(newPath));
            File.Move(source, newPath);

            return true;
        }
    }
}
