using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShortcutSync
{
    public static class Extensions
    {
        public static string GetSafeFileName(this string validName)
        {
            foreach (var c in Path.GetInvalidFileNameChars()) validName = validName.Replace(c.ToString(), "");
            return validName;
        }

        public static string ToHexCode(this Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        public static bool AddMissing<Key, Value>(this IDictionary<Key, Value> dict, Key key, Value value)
        {
            if (dict.ContainsKey(key))
            {
                return false;
            } else
            {
                dict.Add(key, value);
                return true;
            }
        }
    }
}
