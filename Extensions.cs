using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ShortcutSync
{
    public static class Extensions
    {
        public static string GetSafeFileName(this string validName)
        {
            foreach (var c in Path.GetInvalidFileNameChars()) validName = validName.Replace(c.ToString(), "");
            return validName;
        }

        public static string ToHexCode(this Color color, bool includeAlpha = false)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}"
                + (includeAlpha ? color.A.ToString("X2") : string.Empty);
        }

        public static bool AddMissing<Key, Value>(this IDictionary<Key, Value> dict, Key key, Value value)
        {
            if (dict.ContainsKey(key))
            {
                return false;
            }
            else
            {
                dict.Add(key, value);
                return true;
            }
        }

        public static bool IsValidPath(this string path)
        {
            var invalidChars = Path.GetInvalidPathChars();
            return !path.ToCharArray().Any(c => invalidChars.Contains(c));
        }

        public enum PathType { Invalid, File, Directory }
        public static PathType GetPathType(this string path)
        {
            if (!path.IsValidPath()) return PathType.Invalid;
            if (Path.HasExtension(path) && Path.GetFileNameWithoutExtension(path) != string.Empty) return PathType.File;
            return PathType.Directory;
        }

        public static T Copy<T>(this T instance) where T : new()
        {
            if (Object.ReferenceEquals(instance, null))
            {
                return default(T);
            }
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(instance));
        }

        public static bool IsNullOrEmpty(this string str)
        {
            return str == null ? true : str == string.Empty;
        }

        public static float Map(this float value, float from1, float to1, float from2, float to2)
        {
            return (value - from1) / (to1 - from1) * (to2 - from2) + from2;
        }
    }
}
