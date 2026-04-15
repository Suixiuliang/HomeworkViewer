// Copyright (c) 2026 MaxSui 隋修梁. All rights reserved.
// Licensed under the GPL3.0 License. See LICENSE in the project root for license information.
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
            string separator = new string('—', 60);
            string dateLine = $"         {date:yyyy年MM月dd日}         ";
            string titleLine = "          作业          ";

            sb.AppendLine(separator);
            sb.AppendLine(dateLine);
            sb.AppendLine(titleLine);
            sb.AppendLine(separator);

            foreach (var subject in HomeworkData.SubjectNames)
            {
                if (!data.Subjects.ContainsKey(subject)) continue;
                string content = data.Subjects[subject] ?? "";
                sb.AppendLine();
                sb.AppendLine($"【{subject}】");
                string[] lines = content.Split('\n');
                foreach (var line in lines)
                {
                    string trimmed = line.TrimEnd();
                    if (!string.IsNullOrEmpty(trimmed))
                        sb.AppendLine(trimmed);
                }
                sb.AppendLine();
                sb.AppendLine(separator);
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
                subjectsHtml.Append($@"
                <div class='subject'>
                    <div class='subject-name'>{subject}</div>
                    <div class='content'>{content.Replace("\n", "<br/>")}</div>
                </div>");
            }
            string html = TemplateHtml.Replace("{date}", date.ToString("yyyy年MM月dd日"))
                                       .Replace("{subjects}", subjectsHtml.ToString());
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
        }
    }
}