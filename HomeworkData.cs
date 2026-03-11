#nullable disable
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace HomeworkViewer
{
    public class HomeworkData
    {
        public Dictionary<string, string> Subjects { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> ClassReps { get; set; } = new Dictionary<string, string>();   // 未使用但保留
        public Dictionary<string, string> Teachers { get; set; } = new Dictionary<string, string>();    // 未使用但保留
        public Dictionary<string, string> DueTimes { get; set; } = new Dictionary<string, string>();

        // 所有可能的科目（用于兼容）
        public static readonly string[] SubjectNames = { "语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理" };

        public static HomeworkData Load(DateTime date)
        {
            string path = GetFilePath(date);
            if (File.Exists(path))
            {
                string json = File.ReadAllText(path);
                return JsonSerializer.Deserialize<HomeworkData>(json) ?? new HomeworkData();
            }
            return new HomeworkData();
        }

        public void Save(DateTime date)
        {
            string path = GetFilePath(date);
            string dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }

        private static string GetFilePath(DateTime date)
        {
            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(documents, "Homework", date.ToString("yyyy"), date.ToString("MM"), date.ToString("dd") + ".json");
        }

        public static List<DateTime> GetAvailableDates()
        {
            var dates = new List<DateTime>();
            string baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Homework");
            if (!Directory.Exists(baseDir)) return dates;

            foreach (string yearDir in Directory.GetDirectories(baseDir))
            {
                string year = Path.GetFileName(yearDir);
                foreach (string monthDir in Directory.GetDirectories(yearDir))
                {
                    string month = Path.GetFileName(monthDir);
                    foreach (string file in Directory.GetFiles(monthDir, "*.json"))
                    {
                        string day = Path.GetFileNameWithoutExtension(file);
                        if (DateTime.TryParse($"{year}-{month}-{day}", out DateTime dt))
                            dates.Add(dt);
                    }
                }
            }
            dates.Sort((a, b) => b.CompareTo(a));
            return dates;
        }
    }
}