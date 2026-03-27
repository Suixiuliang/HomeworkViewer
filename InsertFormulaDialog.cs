using System.Windows.Forms;

namespace HomeworkViewer
{
    public class InsertFormulaDialog : Form
    {
        private TextBox textBox;
        private Button btnOK;
        private Button btnCancel;

        public string? FormulaText { get; private set; }   // 改为可空类型

        public InsertFormulaDialog()
        {
            this.Text = "插入 LaTeX 公式";
            this.Size = new System.Drawing.Size(400, 150);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label label = new Label
            {
                Text = "请输入 LaTeX 公式（例如：E=mc^2）：",
                Location = new System.Drawing.Point(12, 15),
                AutoSize = true
            };

            textBox = new TextBox
            {
                Location = new System.Drawing.Point(12, 45),
                Size = new System.Drawing.Size(360, 25),
                Font = new System.Drawing.Font("微软雅黑", 10)
            };

            btnOK = new Button
            {
                Text = "确定",
                Location = new System.Drawing.Point(220, 85),
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
                Location = new System.Drawing.Point(305, 85),
                Size = new System.Drawing.Size(75, 23)
            };
            btnCancel.Click += (s, e) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };

            this.Controls.Add(label);
            this.Controls.Add(textBox);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);
        }
    }
}