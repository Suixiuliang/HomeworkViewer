using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace HomeworkViewer
{
    public class InsertFormulaDialog : Form
    {
        private TextBox textBox;
        private Button btnOK;
        private Button btnCancel;
        private Button btnHelp;
        private Label lblInfo;
        private LinkLabel linkHelp;

        public string? FormulaText { get; private set; }

        public InsertFormulaDialog()
        {
            this.Text = "插入 LaTeX 公式";
            this.Size = new System.Drawing.Size(450, 200);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label label = new Label
            {
                Text = "请输入 LaTeX 格式的数学公式（例如：E=mc^2）：",
                Location = new System.Drawing.Point(12, 15),
                AutoSize = true
            };

            textBox = new TextBox
            {
                Location = new System.Drawing.Point(12, 45),
                Size = new System.Drawing.Size(410, 25),
                Font = new System.Drawing.Font("微软雅黑", 10)
            };

            btnOK = new Button
            {
                Text = "确定",
                Location = new System.Drawing.Point(220, 130),
                Size = new System.Drawing.Size(75, 23)
            };
            btnOK.Click += (s, e) =>
            {
                FormulaText = textBox.Text;
                DialogResult = DialogResult.OK;
                Close();
            };

            btnCancel = new Button
            {
                Text = "取消",
                Location = new System.Drawing.Point(305, 130),
                Size = new System.Drawing.Size(75, 23)
            };
            btnCancel.Click += (s, e) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };

            btnHelp = new Button
            {
                Text = "帮助",
                Location = new System.Drawing.Point(135, 130),
                Size = new System.Drawing.Size(75, 23)
            };
            btnHelp.Click += (s, e) =>
            {
                try
                {
                    // 使用 UseShellExecute 确保通过系统默认浏览器打开
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "https://en.wikibooks.org/wiki/LaTeX",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法打开帮助页面：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            lblInfo = new Label
            {
                Text = "💡 提示：使用 Ctrl + I 可快速打开此窗口",
                Location = new System.Drawing.Point(12, 85),
                AutoSize = true,
                ForeColor = System.Drawing.Color.Gray,
                Font = new System.Drawing.Font("微软雅黑", 8)
            };

            linkHelp = new LinkLabel
            {
                Text = "什么是 LaTeX？",
                Location = new System.Drawing.Point(12, 105),
                AutoSize = true,
                LinkColor = System.Drawing.Color.Blue
            };
            linkHelp.LinkClicked += (s, e) =>
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "https://www.latex-project.org/",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法打开链接：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            this.Controls.Add(label);
            this.Controls.Add(textBox);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);
            this.Controls.Add(btnHelp);
            this.Controls.Add(lblInfo);
            this.Controls.Add(linkHelp);
        }
    }
}