using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace HomeworkViewer
{
    public static class FontManager
    {
        private static PrivateFontCollection _privateFonts = new PrivateFontCollection();
        private static bool _customLoaded = false;

        public static void LoadCustomFont(string fontPath)
        {
            if (!File.Exists(fontPath)) return;
            _privateFonts.AddFontFile(fontPath);
            _customLoaded = true;
        }

        public static Font GetFont(string fontFamilyName, float size, FontStyle style = FontStyle.Regular)
        {
            if (!string.IsNullOrEmpty(fontFamilyName))
            {
                try
                {
                    return new Font(fontFamilyName, size, style);
                }
                catch { }
            }
            return new Font("微软雅黑", size, style);
        }

        public static Font GetCustomFont(float size, FontStyle style = FontStyle.Regular)
        {
            if (_customLoaded && _privateFonts.Families.Length > 0)
            {
                return new Font(_privateFonts.Families[0], size, style);
            }
            return GetFont("微软雅黑", size, style);
        }
    }
}