#nullable disable
using System;
using System.Drawing;
using System.Windows.Forms;

namespace HomeworkViewer
{
    public class SettingsForm : Form
    {
        // 侧边栏按钮
        private Button btnBasic;
        private Button btnAppearance;
        private Button btnAbout;

        // 右侧内容面板
        private Panel contentPanel;

        // 基本设置控件
        private ComboBox cmbMode;
        private ComboBox cmbFontSize;

        // 外观设置控件
        private TrackBar trackCardOpacity;
        private NumericUpDown numCardOpacity;
        private TrackBar trackBgOpacity;
        private NumericUpDown numBgOpacity;
        private RadioButton rbBlack;
        private RadioButton rbWhite;
        // 新增：Bar 颜色选择
        private Button btnBarColor;
        private Panel pnlBarColorPreview;
        private ColorDialog colorDialog;

        // 关于控件
        private Label lblVersion;

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
            this.Size = new Size(600, 420);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Panel sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 120,
                BackColor = Color.FromArgb(30, 30, 30)
            };

            btnBasic = CreateSidebarButton("基本设置", 0);
            btnAppearance = CreateSidebarButton("外观", 1);
            btnAbout = CreateSidebarButton("关于", 2);

            sidebar.Controls.Add(btnBasic);
            sidebar.Controls.Add(btnAppearance);
            sidebar.Controls.Add(btnAbout);

            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(45, 45, 48),
                Padding = new Padding(20)
            };

            Panel buttonPanel = new Panel
            {
                Height = 50,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(30, 30, 30)
            };
            btnOK = new Button
            {
                Text = "确定",
                Location = new Point(200, 10),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(290, 10),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnOK.Click += BtnOK_Click;
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);

            this.Controls.Add(contentPanel);
            this.Controls.Add(sidebar);
            this.Controls.Add(buttonPanel);

            CreateBasicPage();
            CreateAppearancePage();
            CreateAboutPage();

            colorDialog = new ColorDialog();
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
            btnBasic.BackColor = Color.FromArgb(30, 30, 30);
            btnAppearance.BackColor = Color.FromArgb(30, 30, 30);
            btnAbout.BackColor = Color.FromArgb(30, 30, 30);

            switch (selectedIndex)
            {
                case 0: btnBasic.BackColor = Color.FromArgb(64, 64, 64); break;
                case 1: btnAppearance.BackColor = Color.FromArgb(64, 64, 64); break;
                case 2: btnAbout.BackColor = Color.FromArgb(64, 64, 64); break;
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
                case 2: ShowAboutPage(); break;
            }
        }

        // ---------- 基本设置页面 ----------
        private Panel basicPanel;
        private void CreateBasicPage()
        {
            basicPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            Label lblMode = new Label
            {
                Text = "展示模式:",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = true
            };
            layout.Controls.Add(lblMode, 0, 0);

            cmbMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 150,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cmbMode.Items.AddRange(new object[] { "大理", "中理", "小理", "大文", "全科" });
            layout.Controls.Add(cmbMode, 1, 0);

            Label lblFontSize = new Label
            {
                Text = "字号大小:",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = true
            };
            layout.Controls.Add(lblFontSize, 0, 1);

            cmbFontSize = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 100,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cmbFontSize.Items.AddRange(new object[] { "小", "中", "大" });
            layout.Controls.Add(cmbFontSize, 1, 1);

            basicPanel.Controls.Add(layout);
        }

        private void ShowBasicPage()
        {
            contentPanel.Controls.Add(basicPanel);
        }

        // ---------- 外观设置页面 ----------
        private Panel appearancePanel;
        private void CreateAppearancePage()
        {
            appearancePanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

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
            Label lblCardOpacity = new Label
            {
                Text = "卡片透明度:",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = true
            };
            layout.Controls.Add(lblCardOpacity, 0, 0);

            Panel cardPanel = new Panel { Height = 30, Width = 250, BackColor = Color.Transparent };
            trackCardOpacity = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                TickFrequency = 10,
                Width = 150,
                Location = new Point(0, 0),
                BackColor = Color.FromArgb(45, 45, 48)
            };
            numCardOpacity = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Width = 50,
                Location = new Point(160, 2),
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White
            };
            trackCardOpacity.ValueChanged += (s, e) => numCardOpacity.Value = trackCardOpacity.Value;
            numCardOpacity.ValueChanged += (s, e) => trackCardOpacity.Value = (int)numCardOpacity.Value;
            cardPanel.Controls.Add(trackCardOpacity);
            cardPanel.Controls.Add(numCardOpacity);
            layout.Controls.Add(cardPanel, 1, 0);

            // 背景透明度
            Label lblBgOpacity = new Label
            {
                Text = "背景透明度:",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = true
            };
            layout.Controls.Add(lblBgOpacity, 0, 1);

            Panel bgPanel = new Panel { Height = 30, Width = 250, BackColor = Color.Transparent };
            trackBgOpacity = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                TickFrequency = 10,
                Width = 150,
                Location = new Point(0, 0),
                BackColor = Color.FromArgb(45, 45, 48)
            };
            numBgOpacity = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Width = 50,
                Location = new Point(160, 2),
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White
            };
            trackBgOpacity.ValueChanged += (s, e) => numBgOpacity.Value = trackBgOpacity.Value;
            numBgOpacity.ValueChanged += (s, e) => trackBgOpacity.Value = (int)numBgOpacity.Value;
            bgPanel.Controls.Add(trackBgOpacity);
            bgPanel.Controls.Add(numBgOpacity);
            layout.Controls.Add(bgPanel, 1, 1);

            // 字体颜色
            Label lblFontColor = new Label
            {
                Text = "字体颜色:",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = true
            };
            layout.Controls.Add(lblFontColor, 0, 2);

            FlowLayoutPanel colorPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Height = 30,
                Width = 200,
                BackColor = Color.Transparent
            };
            rbBlack = new RadioButton
            {
                Text = "黑色",
                Checked = true,
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };
            rbWhite = new RadioButton
            {
                Text = "白色",
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };
            colorPanel.Controls.Add(rbBlack);
            colorPanel.Controls.Add(rbWhite);
            layout.Controls.Add(colorPanel, 1, 2);

            // 新增：顶部 Bar 颜色
            Label lblBarColor = new Label
            {
                Text = "顶部条颜色:",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = true
            };
            layout.Controls.Add(lblBarColor, 0, 3);

            FlowLayoutPanel barColorPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Height = 30,
                Width = 250,
                BackColor = Color.Transparent
            };
            btnBarColor = new Button
            {
                Text = "选择颜色",
                Width = 80,
                Height = 25,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnBarColor.Click += BtnBarColor_Click;

            pnlBarColorPreview = new Panel
            {
                Width = 40,
                Height = 25,
                BackColor = Color.Yellow,
                BorderStyle = BorderStyle.FixedSingle
            };

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
                if (parts.Length == 3)
                {
                    int r = int.Parse(parts[0]);
                    int g = int.Parse(parts[1]);
                    int b = int.Parse(parts[2]);
                    return Color.FromArgb(r, g, b);
                }
            }
            catch { }
            return defaultColor;
        }

        private void ShowAppearancePage()
        {
            contentPanel.Controls.Add(appearancePanel);
        }

        // ---------- 关于页面 ----------
        private Panel aboutPanel;
        private void CreateAboutPage()
        {
            aboutPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            lblVersion = new Label
            {
                Text = "作业展板 版本 1.1.0",
                Font = new Font("微软雅黑", 12, FontStyle.Bold),
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };
            layout.Controls.Add(lblVersion, 0, 0);

            Label lblCopyright = new Label
            {
                Text = "© 2026 保留部分权利",
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };
            layout.Controls.Add(lblCopyright, 0, 1);

            aboutPanel.Controls.Add(layout);
        }

        private void ShowAboutPage()
        {
            contentPanel.Controls.Add(aboutPanel);
        }

        // ---------- 数据加载与保存 ----------
        private void LoadSettings()
        {
            cmbMode.SelectedItem = config.LastMode;
            cmbFontSize.SelectedIndex = config.FontSizeLevel;
            trackCardOpacity.Value = config.CardOpacity;
            trackBgOpacity.Value = config.BackgroundOpacity;
            if (config.FontColorWhite)
                rbWhite.Checked = true;
            else
                rbBlack.Checked = true;

            pnlBarColorPreview.BackColor = ParseColor(config.BarColor, Color.Yellow);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            config.LastMode = cmbMode.SelectedItem?.ToString() ?? "大理";
            config.FontSizeLevel = cmbFontSize.SelectedIndex;
            config.CardOpacity = trackCardOpacity.Value;
            config.BackgroundOpacity = trackBgOpacity.Value;
            config.FontColorWhite = rbWhite.Checked;
            // Bar 颜色已在点击时保存到 config.BarColor，无需额外处理
            config.Save();

            mainForm.ApplySettings(config);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}