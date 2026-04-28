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
        // 基本设置
        public string LastMode { get; set; } = "大理";
        public int FontSizeLevel { get; set; } = 1;
        public int CardOpacity { get; set; } = 15;
        public bool FontColorWhite { get; set; } = false;
        public string BarColor { get; set; } = "255,255,0";

        // 晚修相关
        public int EveningClassCount { get; set; } = 3;
        public List<EveningClassTime> EveningClassTimes { get; set; } = new List<EveningClassTime>();

        // 滚动速度 (px/s)
        public int ScrollSpeed { get; set; } = 30;

        // 自动翻页间隔 (秒)
        public int AutoPageInterval { get; set; } = 10;

        // 科代表名单（未完全使用）
        public Dictionary<string, string> ClassReps { get; set; } = new Dictionary<string, string>();

        // 背景效果 (Mica / Acrylic / Aero)
        public string BackgroundEffect { get; set; } = "Mica";

        // 更新标记
        public int UpdatePending { get; set; } = 0;

        // 显示设置
        public bool ShowDueTime { get; set; } = true;
        public bool ShowMouseGlow { get; set; } = true;

        // 背景图片
        public bool UseBackgroundImage { get; set; } = false;
        public string BackgroundImagePath { get; set; } = "";

        // 字体
        public string FontFamily { get; set; } = "微软雅黑";
        public bool IsCustomFont { get; set; } = false;

        // 卡片行列尺寸（0 表示自动）
        public int RowHeight { get; set; } = 0;
        public int ColumnWidth { get; set; } = 0;

        // 默认导出格式
        public string ExportFormat { get; set; } = "txt";

        // 集控管理
        public bool MgmtEnabled { get; set; } = false;
        public string MgmtManifestUrl { get; set; } = "";
        public Dictionary<string, int> MgmtVersions { get; set; } = new();
        public bool MgmtForceRemote { get; set; } = false;
        public string OrganizationName { get; set; } = "";

        // ========== 新增功能 ==========
        // 自定义卡片科目列表（顺序即卡片显示顺序，每页 3x2）
        public List<string> CustomSubjects { get; set; } = new List<string>();

        // 周末作业延续：周五的作业自动显示在周六、周日
        public bool ExtendFridayHomeworkToWeekend { get; set; } = true;

        // ========== 已废弃但保留兼容的属性（不再使用） ==========
        [Obsolete("已废弃，不再使用")]
        public int BackgroundOpacity { get; set; } = 12;

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

                    // 确保晚修时间段列表与节数一致
                    if (config.EveningClassTimes == null)
                        config.EveningClassTimes = new List<EveningClassTime>();
                    while (config.EveningClassTimes.Count < config.EveningClassCount)
                        config.EveningClassTimes.Add(new EveningClassTime { Start = "19:00", End = "19:50" });
                    if (config.EveningClassTimes.Count > config.EveningClassCount)
                        config.EveningClassTimes = config.EveningClassTimes.GetRange(0, config.EveningClassCount);

                    // 确保 ClassReps 不为空
                    if (config.ClassReps == null)
                        config.ClassReps = new Dictionary<string, string>();

                    // 确保 CustomSubjects 不为空
                    if (config.CustomSubjects == null || config.CustomSubjects.Count == 0)
                    {
                        config.CustomSubjects = new List<string> { "语文", "数学", "英语", "物理", "化学", "生物" };
                    }

                    // 确保 AutoPageInterval 有默认值
                    if (config.AutoPageInterval <= 0) config.AutoPageInterval = 10;

                    return config;
                }
                catch { }
            }

            // 返回默认配置
            var defaultConfig = new AppConfig();
            defaultConfig.EveningClassTimes.Clear();
            for (int i = 0; i < defaultConfig.EveningClassCount; i++)
                defaultConfig.EveningClassTimes.Add(new EveningClassTime { Start = "19:00", End = "19:50" });
            defaultConfig.ClassReps = new Dictionary<string, string>();
            defaultConfig.CustomSubjects = new List<string> { "语文", "数学", "英语", "物理", "化学", "生物" };
            defaultConfig.AutoPageInterval = 10;
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