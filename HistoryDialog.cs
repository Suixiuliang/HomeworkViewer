#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace HomeworkViewer
{
    public class HistoryDialog : Form
    {
        private ComboBox cmbYear;
        private ComboBox cmbMonth;
        private ComboBox cmbDay;
        private Button btnOk;
        private Button btnCancel;

        private List<DateTime> availableDates;
        private DateTime? selectedDate;

        public DateTime? SelectedDate => selectedDate;

        public HistoryDialog()
        {
            InitializeComponent();
            LoadData();
            ApplyTransparentEffect();
        }

        private void InitializeComponent()
        {
            this.Text = "选择历史日期";
            this.Size = new Size(300, 250);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblYear = new Label { Text = "年份:", Location = new Point(20, 20), AutoSize = true };
            cmbYear = new ComboBox { Location = new Point(80, 17), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbYear.SelectedIndexChanged += (s, e) => UpdateMonths();

            Label lblMonth = new Label { Text = "月份:", Location = new Point(20, 50), AutoSize = true };
            cmbMonth = new ComboBox { Location = new Point(80, 47), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbMonth.SelectedIndexChanged += (s, e) => UpdateDays();

            Label lblDay = new Label { Text = "日期:", Location = new Point(20, 80), AutoSize = true };
            cmbDay = new ComboBox { Location = new Point(80, 77), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };

            btnOk = new Button { Text = "确定", Location = new Point(80, 120), Width = 70 };
            btnOk.Click += (s, e) => ConfirmSelection();

            btnCancel = new Button { Text = "取消", Location = new Point(160, 120), Width = 70 };
            btnCancel.Click += (s, e) => { selectedDate = null; DialogResult = DialogResult.Cancel; Close(); };

            Controls.AddRange(new Control[] { lblYear, cmbYear, lblMonth, cmbMonth, lblDay, cmbDay, btnOk, btnCancel });
        }

        private void ApplyTransparentEffect()
        {
            this.Opacity = 0.95;
            this.BackColor = Color.Black;
            this.ForeColor = Color.White;

            foreach (Control c in this.Controls)
            {
                if (c is Label label)
                {
                    label.BackColor = Color.Transparent;
                    label.ForeColor = Color.White;
                }
                else if (c is ComboBox combo)
                {
                    combo.BackColor = Color.FromArgb(64, 64, 64);
                    combo.ForeColor = Color.White;
                    combo.FlatStyle = FlatStyle.Flat;
                }
                else if (c is Button btn)
                {
                    btn.BackColor = Color.FromArgb(64, 64, 64);
                    btn.ForeColor = Color.White;
                    btn.FlatStyle = FlatStyle.Flat;
                }
            }
        }

        private void LoadData()
        {
            availableDates = HomeworkData.GetAvailableDates();
            cmbYear.Items.AddRange(availableDates.Select(d => d.Year.ToString()).Distinct().ToArray());
            if (cmbYear.Items.Count > 0) cmbYear.SelectedIndex = 0;
        }

        private void UpdateMonths()
        {
            if (cmbYear.SelectedItem == null) return;
            int year = int.Parse(cmbYear.SelectedItem.ToString());
            var months = availableDates.Where(d => d.Year == year).Select(d => d.Month).Distinct().OrderByDescending(m => m);
            cmbMonth.Items.Clear();
            foreach (var m in months) cmbMonth.Items.Add(m.ToString("00"));
            if (cmbMonth.Items.Count > 0) cmbMonth.SelectedIndex = 0;
        }

        private void UpdateDays()
        {
            if (cmbYear.SelectedItem == null || cmbMonth.SelectedItem == null) return;
            int year = int.Parse(cmbYear.SelectedItem.ToString());
            int month = int.Parse(cmbMonth.SelectedItem.ToString());
            var days = availableDates.Where(d => d.Year == year && d.Month == month).Select(d => d.Day).Distinct().OrderByDescending(d => d);
            cmbDay.Items.Clear();
            foreach (var d in days) cmbDay.Items.Add(d.ToString("00"));
            if (cmbDay.Items.Count > 0) cmbDay.SelectedIndex = 0;
        }

        private void ConfirmSelection()
        {
            if (cmbYear.SelectedItem != null && cmbMonth.SelectedItem != null && cmbDay.SelectedItem != null)
            {
                int year = int.Parse(cmbYear.SelectedItem.ToString());
                int month = int.Parse(cmbMonth.SelectedItem.ToString());
                int day = int.Parse(cmbDay.SelectedItem.ToString());
                selectedDate = new DateTime(year, month, day);
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("请选择完整的日期", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}