#nullable disable
using System;
using System.Drawing;
using System.IO;
using System.Text.Json;

namespace HomeworkViewer
{
    public class AppConfig
    {
        public string LastMode { get; set; } = "大理";
        public int FontSizeLevel { get; set; } = 1;
        public int CardOpacity { get; set; } = 15;
        public int BackgroundOpacity { get; set; } = 12;
        public bool FontColorWhite { get; set; } = false;
        public string BarColor { get; set; } = "255,255,0"; // 默认黄色

        private static string GetConfigPath()
        {
            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(documents, "Homework", "config.json");
        }

        public static AppConfig Load()
        {
            string path = GetConfigPath();
            if (File.Exists(path))
            {
                try
                {
                    string json = File.ReadAllText(path);
                    return JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                }
                catch
                {
                    return new AppConfig();
                }
            }
            return new AppConfig();
        }

        public void Save()
        {
            string path = GetConfigPath();
            string dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
    }
}