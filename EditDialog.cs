using System;
using System.Windows.Forms;

namespace HomeworkViewer
{
    // 此文件已不再使用，因为改为内联编辑。保留仅作参考。
    public class EditDialog : Form
    {
        private TextBox textBox;
        private Button saveButton;
        private Button cancelButton;

        public string EditedText { get; private set; }

        public EditDialog(string subject, string initialText)
        {
            Text = $"编辑 {subject} 作业";
            Size = new System.Drawing.Size(600, 400);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            Label infoLabel = new Label
            {
                Text = "输入作业内容，保存后自动关闭",
                AutoSize = true,
                Location = new System.Drawing.Point(12, 9)
            };

            textBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(560, 300),
                Text = initialText,
                Font = new System.Drawing.Font("微软雅黑", 11)
            };

            saveButton = new Button
            {
                Text = "保存",
                Location = new System.Drawing.Point(400, 345),
                Size = new System.Drawing.Size(80, 25)
            };
            saveButton.Click += (s, e) => { EditedText = textBox.Text; DialogResult = DialogResult.OK; Close(); };

            cancelButton = new Button
            {
                Text = "取消",
                Location = new System.Drawing.Point(490, 345),
                Size = new System.Drawing.Size(80, 25)
            };
            cancelButton.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            Controls.Add(infoLabel);
            Controls.Add(textBox);
            Controls.Add(saveButton);
            Controls.Add(cancelButton);
        }
    }
}