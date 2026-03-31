#nullable disable
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace HomeworkViewer
{
    public class EveningClassTime
    {
        public string Start { get; set; }
        public string End { get; set; }
    }

    public class AppConfig
    {
        public string LastMode { get; set; } = "大理";
        public int FontSizeLevel { get; set; } = 1;
        public int CardOpacity { get; set; } = 15;
        public int BackgroundOpacity { get; set; } = 12;
        public bool FontColorWhite { get; set; } = false;
        public string BarColor { get; set; } = "255,255,0";

        public int EveningClassCount { get; set; } = 3;
        public List<EveningClassTime> EveningClassTimes { get; set; } = new List<EveningClassTime>();

        public int ScrollSpeed { get; set; } = 30;

        public Dictionary<string, string> ClassReps { get; set; } = new Dictionary<string, string>();

        public string BackgroundEffect { get; set; } = "Mica";

        public int UpdatePending { get; set; } = 0;

        // Markdown 渲染开关
        public bool EnableMarkdown { get; set; } = false;

        // 字体设置
        public string FontFamily { get; set; } = "微软雅黑";
        public bool IsCustomFont { get; set; } = false;

        // 卡片尺寸调整（0 表示自动）
        public int RowHeight { get; set; } = 0;
        public int ColumnWidth { get; set; } = 0;

        // 导出默认格式
        public string ExportFormat { get; set; } = "txt";

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
                    var config = JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                    if (config.EveningClassTimes == null)
                        config.EveningClassTimes = new List<EveningClassTime>();
                    while (config.EveningClassTimes.Count < config.EveningClassCount)
                        config.EveningClassTimes.Add(new EveningClassTime { Start = "19:00", End = "19:50" });
                    if (config.EveningClassTimes.Count > config.EveningClassCount)
                        config.EveningClassTimes = config.EveningClassTimes.GetRange(0, config.EveningClassCount);
                    if (config.ClassReps == null)
                        config.ClassReps = new Dictionary<string, string>();
                    return config;
                }
                catch { }
            }
            var defaultConfig = new AppConfig();
            defaultConfig.EveningClassTimes.Clear();
            for (int i = 0; i < defaultConfig.EveningClassCount; i++)
                defaultConfig.EveningClassTimes.Add(new EveningClassTime { Start = "19:00", End = "19:50" });
            defaultConfig.ClassReps = new Dictionary<string, string>();
            return defaultConfig;
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