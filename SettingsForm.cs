#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace HomeworkViewer
{
    public class SettingsForm : Form
    {
        private Button btnBasic;
        private Button btnAppearance;
        private Button btnTimeSlot;
        private Button btnReps;
        private Button btnAbout;
        private Panel contentPanel;

        // 基本设置
        private ComboBox cmbMode;
        private ComboBox cmbFontSize;
        private NumericUpDown numScrollSpeed;

        // 外观设置
        private TrackBar trackCardOpacity;
        private NumericUpDown numCardOpacity;
        private TrackBar trackBgOpacity;
        private NumericUpDown numBgOpacity;
        private RadioButton rbBlack;
        private RadioButton rbWhite;
        private Button btnBarColor;
        private Panel pnlBarColorPreview;
        private ColorDialog colorDialog;

        // 时间段设置
        private NumericUpDown numEveningCount;
        private FlowLayoutPanel eveningPanel;
        private List<DateTimePicker> startPickers = new List<DateTimePicker>();
        private List<DateTimePicker> endPickers = new List<DateTimePicker>();
        private Button btnTestFlash;
        private Button btnStopFlash;

        // 科代表设置
        private FlowLayoutPanel repsPanel;
        private List<TextBox> repTextboxes = new List<TextBox>();
        private List<string> currentSubjectsForReps = new List<string>();

        private Button btnOK;
        private Button btnCancel;
        private AppConfig config;
        private HomeworkViewer mainForm;

        private int _selectedPage = 0;

        public SettingsForm(HomeworkViewer main)
        {
            mainForm = main;
            config = AppConfig.Load();

            InitializeComponent();
            LoadSettings();
            ShowPage(0);
            this.Opacity = 0.95;
            this.BackColor = Color.Black;
            this.ForeColor = Color.White;
        }

        private void InitializeComponent()
        {
            this.Text = "设置";
            this.Size = new Size(650, 650);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Panel sidebar = new Panel { Dock = DockStyle.Left, Width = 120, BackColor = Color.FromArgb(30, 30, 30) };
            
            // 创建按钮，Tag 仍保持原有页面索引
            btnAbout = CreateSidebarButton("关于", 4);
            btnReps = CreateSidebarButton("科代表", 3);
            btnTimeSlot = CreateSidebarButton("时间段", 2);
            btnAppearance = CreateSidebarButton("外观", 1);
            btnBasic = CreateSidebarButton("基本设置", 0);

            // 按从下到上的顺序添加（后添加的会出现在顶部）
            // 我们希望基本设置在最上面，所以最后添加 btnBasic
            sidebar.Controls.Add(btnAbout);      // 最下面
            sidebar.Controls.Add(btnReps);
            sidebar.Controls.Add(btnTimeSlot);
            sidebar.Controls.Add(btnAppearance);
            sidebar.Controls.Add(btnBasic);      // 最上面

            contentPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(45, 45, 48), Padding = new Padding(20), AutoScroll = true };

            Panel buttonPanel = new Panel { Height = 50, Dock = DockStyle.Bottom, BackColor = Color.FromArgb(30, 30, 30) };
            btnOK = new Button { Text = "确定", Location = new Point(200, 10), Size = new Size(80, 30), BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnCancel = new Button { Text = "取消", Location = new Point(290, 10), Size = new Size(80, 30), BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnOK.Click += BtnOK_Click;
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;
            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);

            this.Controls.Add(contentPanel);
            this.Controls.Add(sidebar);
            this.Controls.Add(buttonPanel);

            colorDialog = new ColorDialog();

            CreateBasicPage();
            CreateAppearancePage();
            CreateTimeSlotPage();
            CreateRepsPage();
            CreateAboutPage();
        }

        private Button CreateSidebarButton(string text, int tag)
        {
            Button btn = new Button
            {
                Text = text,
                Dock = DockStyle.Top,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0),
                FlatAppearance = { BorderSize = 0 }
            };
            btn.Tag = tag;
            btn.Click += SidebarButton_Click;
            return btn;
        }

        private void SidebarButton_Click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null)
            {
                int pageIndex = (int)btn.Tag;
                ShowPage(pageIndex);
                UpdateSidebarSelection(pageIndex);
            }
        }

        private void UpdateSidebarSelection(int selectedIndex)
        {
            btnBasic.BackColor = btnAppearance.BackColor = btnTimeSlot.BackColor = btnReps.BackColor = btnAbout.BackColor = Color.FromArgb(30, 30, 30);
            switch (selectedIndex)
            {
                case 0: btnBasic.BackColor = Color.FromArgb(64, 64, 64); break;
                case 1: btnAppearance.BackColor = Color.FromArgb(64, 64, 64); break;
                case 2: btnTimeSlot.BackColor = Color.FromArgb(64, 64, 64); break;
                case 3: btnReps.BackColor = Color.FromArgb(64, 64, 64); break;
                case 4: btnAbout.BackColor = Color.FromArgb(64, 64, 64); break;
            }
        }

        private void ShowPage(int pageIndex)
        {
            _selectedPage = pageIndex;
            contentPanel.Controls.Clear();
            switch (pageIndex)
            {
                case 0: ShowBasicPage(); break;
                case 1: ShowAppearancePage(); break;
                case 2: ShowTimeSlotPage(); break;
                case 3: ShowRepsPage(); break;
                case 4: ShowAboutPage(); break;
            }
        }

        // ---------- 基本设置 ----------
        private Panel basicPanel;
        private void CreateBasicPage()
        {
            basicPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 4,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            Label lblMode = new Label { Text = "展示模式:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblMode, 0, 0);
            cmbMode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            cmbMode.Items.AddRange(new object[] { "大理", "中理", "小理", "大文", "全科" });
            layout.Controls.Add(cmbMode, 1, 0);

            Label lblFontSize = new Label { Text = "字号大小:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblFontSize, 0, 1);
            cmbFontSize = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            cmbFontSize.Items.AddRange(new object[] { "小", "中", "大" });
            layout.Controls.Add(cmbFontSize, 1, 1);

            Label lblScrollSpeed = new Label { Text = "滚动速度(px/s):", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblScrollSpeed, 0, 2);
            numScrollSpeed = new NumericUpDown { Minimum = 0, Maximum = 200, Value = config.ScrollSpeed, Width = 60, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
            layout.Controls.Add(numScrollSpeed, 1, 2);

            basicPanel.Controls.Add(layout);
        }
        private void ShowBasicPage() => contentPanel.Controls.Add(basicPanel);

        // ---------- 外观设置 ----------
        private Panel appearancePanel;
        private void CreateAppearancePage()
        {
            appearancePanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 5,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // 卡片透明度
            Label lblCardOpacity = new Label { Text = "卡片透明度:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblCardOpacity, 0, 0);
            Panel cardPanel = new Panel { Height = 30, Width = 250, BackColor = Color.Transparent };
            trackCardOpacity = new TrackBar { Minimum = 0, Maximum = 100, TickFrequency = 10, Width = 150, Location = new Point(0, 0), BackColor = Color.FromArgb(45, 45, 48) };
            numCardOpacity = new NumericUpDown { Minimum = 0, Maximum = 100, Width = 50, Location = new Point(160, 2), BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
            trackCardOpacity.ValueChanged += (s, e) => numCardOpacity.Value = trackCardOpacity.Value;
            numCardOpacity.ValueChanged += (s, e) => trackCardOpacity.Value = (int)numCardOpacity.Value;
            cardPanel.Controls.Add(trackCardOpacity);
            cardPanel.Controls.Add(numCardOpacity);
            layout.Controls.Add(cardPanel, 1, 0);

            // 背景透明度
            Label lblBgOpacity = new Label { Text = "背景透明度:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblBgOpacity, 0, 1);
            Panel bgPanel = new Panel { Height = 30, Width = 250, BackColor = Color.Transparent };
            trackBgOpacity = new TrackBar { Minimum = 0, Maximum = 100, TickFrequency = 10, Width = 150, Location = new Point(0, 0), BackColor = Color.FromArgb(45, 45, 48) };
            numBgOpacity = new NumericUpDown { Minimum = 0, Maximum = 100, Width = 50, Location = new Point(160, 2), BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
            trackBgOpacity.ValueChanged += (s, e) => numBgOpacity.Value = trackBgOpacity.Value;
            numBgOpacity.ValueChanged += (s, e) => trackBgOpacity.Value = (int)numBgOpacity.Value;
            bgPanel.Controls.Add(trackBgOpacity);
            bgPanel.Controls.Add(numBgOpacity);
            layout.Controls.Add(bgPanel, 1, 1);

            // 字体颜色
            Label lblFontColor = new Label { Text = "字体颜色:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblFontColor, 0, 2);
            FlowLayoutPanel colorPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Height = 30, Width = 200, BackColor = Color.Transparent };
            rbBlack = new RadioButton { Text = "黑色", Checked = true, AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            rbWhite = new RadioButton { Text = "白色", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            colorPanel.Controls.Add(rbBlack);
            colorPanel.Controls.Add(rbWhite);
            layout.Controls.Add(colorPanel, 1, 2);

            // Bar 颜色
            Label lblBarColor = new Label { Text = "顶部条颜色:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblBarColor, 0, 3);
            FlowLayoutPanel barColorPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Height = 30, Width = 250, BackColor = Color.Transparent };
            btnBarColor = new Button { Text = "选择颜色", Width = 80, Height = 25, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnBarColor.Click += BtnBarColor_Click;
            pnlBarColorPreview = new Panel { Width = 40, Height = 25, BackColor = Color.Yellow, BorderStyle = BorderStyle.FixedSingle };
            barColorPanel.Controls.Add(btnBarColor);
            barColorPanel.Controls.Add(pnlBarColorPreview);
            layout.Controls.Add(barColorPanel, 1, 3);

            appearancePanel.Controls.Add(layout);
        }
        private void BtnBarColor_Click(object sender, EventArgs e)
        {
            Color currentColor = ParseColor(config.BarColor, Color.Yellow);
            colorDialog.Color = currentColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                pnlBarColorPreview.BackColor = colorDialog.Color;
                config.BarColor = $"{colorDialog.Color.R},{colorDialog.Color.G},{colorDialog.Color.B}";
            }
        }
        private Color ParseColor(string rgbString, Color defaultColor)
        {
            try
            {
                string[] parts = rgbString.Split(',');
                if (parts.Length == 3) return Color.FromArgb(int.Parse(parts[0]), int.Parse(parts[1]), int.Parse(parts[2]));
            }
            catch { }
            return defaultColor;
        }
        private void ShowAppearancePage() => contentPanel.Controls.Add(appearancePanel);

        // ---------- 时间段设置 ----------
        private Panel timeSlotPanel;
        private void CreateTimeSlotPage()
        {
            timeSlotPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent, AutoScroll = true };
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            TableLayoutPanel countLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            countLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            countLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Label lblCount = new Label { Text = "晚修节数:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            countLayout.Controls.Add(lblCount, 0, 0);
            numEveningCount = new NumericUpDown { Minimum = 1, Maximum = 6, Value = config.EveningClassCount, Width = 60, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
            numEveningCount.ValueChanged += NumEveningCount_ValueChanged;
            countLayout.Controls.Add(numEveningCount, 1, 0);
            mainLayout.Controls.Add(countLayout, 0, 0);

            eveningPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(10, 10, 0, 10)
            };
            mainLayout.Controls.Add(eveningPanel, 0, 1);

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                BackColor = Color.Transparent,
                Padding = new Padding(10, 10, 0, 0)
            };
            btnTestFlash = new Button { Text = "测试闪烁", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnTestFlash.Click += BtnTestFlash_Click;

            btnStopFlash = new Button { Text = "停止闪烁", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            btnStopFlash.Click += BtnStopFlash_Click;

            buttonPanel.Controls.Add(btnTestFlash);
            buttonPanel.Controls.Add(btnStopFlash);
            mainLayout.Controls.Add(buttonPanel, 0, 2);

            timeSlotPanel.Controls.Add(mainLayout);
            UpdateEveningEditors();
        }
        private void BtnTestFlash_Click(object sender, EventArgs e)
        {
            mainForm.StartDebugFlashing();
            MessageBox.Show("所有卡片将闪烁5分钟", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void BtnStopFlash_Click(object sender, EventArgs e)
        {
            mainForm.StopDebugFlashing();
            MessageBox.Show("闪烁已停止", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void NumEveningCount_ValueChanged(object sender, EventArgs e)
        {
            config.EveningClassCount = (int)numEveningCount.Value;
            UpdateEveningEditors();
        }
        private void UpdateEveningEditors()
        {
            eveningPanel.Controls.Clear();
            startPickers.Clear();
            endPickers.Clear();
            while (config.EveningClassTimes.Count < config.EveningClassCount)
                config.EveningClassTimes.Add(new EveningClassTime { Start = "19:00", End = "19:50" });
            if (config.EveningClassTimes.Count > config.EveningClassCount)
                config.EveningClassTimes = config.EveningClassTimes.GetRange(0, config.EveningClassCount);
            for (int i = 0; i < config.EveningClassCount; i++)
            {
                var time = config.EveningClassTimes[i];
                Panel row = new Panel { Height = 35, Width = 450, BackColor = Color.Transparent };
                Label lbl = new Label { Text = $"晚修{i + 1}:", Location = new Point(0, 8), AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
                DateTimePicker startPicker = new DateTimePicker
                {
                    Format = DateTimePickerFormat.Custom,
                    CustomFormat = "HH:mm",
                    ShowUpDown = true,
                    Location = new Point(60, 5),
                    Width = 80,
                    Value = DateTime.TryParse(time.Start, out DateTime start) ? start : DateTime.Today.AddHours(19),
                    BackColor = Color.FromArgb(64, 64, 64),
                    ForeColor = Color.White
                };
                startPickers.Add(startPicker);
                Label lblTo = new Label { Text = "—", Location = new Point(150, 8), AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
                DateTimePicker endPicker = new DateTimePicker
                {
                    Format = DateTimePickerFormat.Custom,
                    CustomFormat = "HH:mm",
                    ShowUpDown = true,
                    Location = new Point(170, 5),
                    Width = 80,
                    Value = DateTime.TryParse(time.End, out DateTime end) ? end : DateTime.Today.AddHours(19).AddMinutes(50),
                    BackColor = Color.FromArgb(64, 64, 64),
                    ForeColor = Color.White
                };
                endPickers.Add(endPicker);
                Panel previewBar = new Panel { Location = new Point(270, 8), Size = new Size(150, 20), BackColor = Color.FromArgb(100, 100, 100), BorderStyle = BorderStyle.FixedSingle };
                row.Controls.Add(lbl);
                row.Controls.Add(startPicker);
                row.Controls.Add(lblTo);
                row.Controls.Add(endPicker);
                row.Controls.Add(previewBar);
                eveningPanel.Controls.Add(row);
            }
        }
        private void ShowTimeSlotPage() => contentPanel.Controls.Add(timeSlotPanel);

        // ---------- 科代表设置 ----------
        private void CreateRepsPage()
        {
            repsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                BackColor = Color.Transparent,
                Padding = new Padding(10)
            };

            // 获取当前模式的科目列表
            var subjectModes = new Dictionary<string, string[]>
            {
                {"大理", new[]{"语文","数学","英语","物理","化学","生物"}},
                {"中理", new[]{"语文","数学","英语","物理","化学","地理"}},
                {"小理", new[]{"语文","数学","英语","物理","化学","政治"}},
                {"大文", new[]{"语文","数学","英语","政治","历史","地理"}},
                {"全科", new[]{"语文","数学","英语","物理","化学","生物","政治","历史","地理"}}
            };
            string[] subjects = subjectModes[config.LastMode];
            currentSubjectsForReps = new List<string>(subjects);

            repTextboxes.Clear();

            foreach (string subject in subjects)
            {
                Panel row = new Panel { Height = 35, Width = 450, BackColor = Color.Transparent };
                Label lbl = new Label { Text = $"{subject}:", Location = new Point(0, 8), AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
                TextBox txt = new TextBox
                {
                    Location = new Point(80, 5),
                    Width = 200,
                    BackColor = Color.FromArgb(64, 64, 64),
                    ForeColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle
                };
                if (config.ClassReps.ContainsKey(subject))
                    txt.Text = config.ClassReps[subject];
                repTextboxes.Add(txt);
                row.Controls.Add(lbl);
                row.Controls.Add(txt);
                repsPanel.Controls.Add(row);
            }
        }
        private void ShowRepsPage() => contentPanel.Controls.Add(repsPanel);

        // ---------- 关于页面 ----------
        private Panel aboutPanel;
        private void CreateAboutPage()
        {
            aboutPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            Label lblVersion = new Label { Text = "作业展板 版本 1.2.0", Font = new Font("微软雅黑", 12, FontStyle.Bold), AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            layout.Controls.Add(lblVersion, 0, 0);
            Label lblAuthor = new Label { Text = "\n作者: MaxSui 隋修梁", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent, Font = new Font("微软雅黑", 10) };
            layout.Controls.Add(lblAuthor, 0, 1);
            Label lblCopyright = new Label { Text = "\n© 2026 MaxSui 保留部分权利 \n本软件遵循GNU General Public License 3.0开源协议", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            layout.Controls.Add(lblCopyright, 0, 2);
            aboutPanel.Controls.Add(layout);
        }
        private void ShowAboutPage() => contentPanel.Controls.Add(aboutPanel);

        // ---------- 加载与保存 ----------
        private void LoadSettings()
        {
            cmbMode.SelectedItem = config.LastMode;
            cmbFontSize.SelectedIndex = config.FontSizeLevel;
            trackCardOpacity.Value = config.CardOpacity;
            trackBgOpacity.Value = config.BackgroundOpacity;
            if (config.FontColorWhite) rbWhite.Checked = true; else rbBlack.Checked = true;
            pnlBarColorPreview.BackColor = ParseColor(config.BarColor, Color.Yellow);
            numScrollSpeed.Value = config.ScrollSpeed;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            config.LastMode = cmbMode.SelectedItem?.ToString() ?? "大理";
            config.FontSizeLevel = cmbFontSize.SelectedIndex;
            config.CardOpacity = trackCardOpacity.Value;
            config.BackgroundOpacity = trackBgOpacity.Value;
            config.FontColorWhite = rbWhite.Checked;
            config.ScrollSpeed = (int)numScrollSpeed.Value;

            config.EveningClassCount = (int)numEveningCount.Value;
            config.EveningClassTimes.Clear();
            for (int i = 0; i < startPickers.Count; i++)
                config.EveningClassTimes.Add(new EveningClassTime { Start = startPickers[i].Value.ToString("HH:mm"), End = endPickers[i].Value.ToString("HH:mm") });

            config.ClassReps.Clear();
            for (int i = 0; i < currentSubjectsForReps.Count && i < repTextboxes.Count; i++)
            {
                string subject = currentSubjectsForReps[i];
                string rep = repTextboxes[i].Text.Trim();
                if (!string.IsNullOrEmpty(rep))
                    config.ClassReps[subject] = rep;
            }

            config.Save();
            mainForm.ApplySettings(config);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}