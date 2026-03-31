using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace HomeworkViewer
{
    public static class ExportHelper
    {
        private static readonly string TemplateHtml = @"<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>作业展板导出</title>
    <style>
        body { font-family: 'Microsoft YaHei', sans-serif; margin: 20px; }
        .date { text-align: center; font-size: 20px; margin-bottom: 20px; }
        .subject { margin-bottom: 20px; border-left: 4px solid #ffcc00; padding-left: 15px; }
        .subject-name { font-size: 18px; font-weight: bold; color: #333; }
        .content { margin-top: 8px; white-space: pre-wrap; }
        .unsubmitted { color: red; margin-top: 5px; font-style: italic; }
    </style>
</head>
<body>
    <div class=""date"">{date}</div>
    {subjects}
</body>
</html>";

        public static void ExportToTxt(HomeworkData data, DateTime date, string filePath, bool includeUnsubmitted)
        {
            var sb = new StringBuilder();
            string dateLine = $"=== {date:yyyy年MM月dd日} ===";
            sb.AppendLine(dateLine);

            foreach (var subject in HomeworkData.SubjectNames)
            {
                if (!data.Subjects.ContainsKey(subject)) continue;
                string content = data.Subjects[subject] ?? "";
                string unsubmitted = includeUnsubmitted && data.Unsubmitted.ContainsKey(subject) ? data.Unsubmitted[subject] : null;
                sb.AppendLine();
                sb.AppendLine($"【{subject}】");
                if (!string.IsNullOrWhiteSpace(content))
                {
                    string[] lines = content.Split('\n');
                    foreach (var line in lines)
                    {
                        string trimmed = line.TrimEnd();
                        if (!string.IsNullOrEmpty(trimmed))
                            sb.AppendLine(trimmed);
                    }
                }
                if (!string.IsNullOrEmpty(unsubmitted))
                {
                    sb.AppendLine($"未交：{unsubmitted}");
                }
            }
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static void ExportToPdf(string txtPath, string pdfPath)
        {
            MessageBox.Show("PDF 导出功能需要安装 iTextSharp，暂未启用。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            File.Copy(txtPath, pdfPath, true);
        }

        public static void ExportToJpg(Form form, string jpgPath)
        {
            using (var bmp = new Bitmap(form.ClientSize.Width, form.ClientSize.Height))
            {
                form.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                bmp.Save(jpgPath, ImageFormat.Jpeg);
            }
        }

        public static void ExportToHtml(HomeworkData data, DateTime date, string htmlPath, bool includeUnsubmitted)
        {
            var subjectsHtml = new StringBuilder();
            foreach (var subject in HomeworkData.SubjectNames)
            {
                if (!data.Subjects.ContainsKey(subject)) continue;
                string content = data.Subjects[subject] ?? "";
                string unsubmitted = includeUnsubmitted && data.Unsubmitted.ContainsKey(subject) ? $"<div class='unsubmitted'>未交：{data.Unsubmitted[subject]}</div>" : "";
                subjectsHtml.Append($@"
                <div class='subject'>
                    <div class='subject-name'>{subject}</div>
                    <div class='content'>{content.Replace("\n", "<br/>")}</div>
                    {unsubmitted}
                </div>");
            }
            string html = TemplateHtml.Replace("{date}", date.ToString("yyyy年MM月dd日"))
                                       .Replace("{subjects}", subjectsHtml.ToString());
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
        }
    }
}