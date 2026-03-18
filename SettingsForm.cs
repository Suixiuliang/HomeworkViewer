#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Reflection;
using System.Diagnostics;
using System.Linq;

namespace HomeworkViewer
{
    public class RemoteVersionInfo
    {
        public string Version { get; set; }
        public string ReleaseType { get; set; }
        public string IsMandatory { get; set; }
    }

    public partial class SettingsForm : Form
    {
        private Button btnBasic;
        private Button btnAppearance;
        private Button btnTimeSlot;
        private Button btnReps;
        private Button btnAbout;
        private Panel contentPanel;

        private ComboBox cmbMode;
        private ComboBox cmbFontSize;
        private NumericUpDown numScrollSpeed;

        private TrackBar trackCardOpacity;
        private NumericUpDown numCardOpacity;
        private TrackBar trackBgOpacity;
        private NumericUpDown numBgOpacity;
        private RadioButton rbBlack;
        private RadioButton rbWhite;
        private Button btnBarColor;
        private Panel pnlBarColorPreview;
        private ColorDialog colorDialog;

        private NumericUpDown numEveningCount;
        private FlowLayoutPanel eveningPanel;
        private List<DateTimePicker> startPickers = new List<DateTimePicker>();
        private List<DateTimePicker> endPickers = new List<DateTimePicker>();
        private Button btnTestFlash;
        private Button btnStopFlash;

        private FlowLayoutPanel repsPanel;
        private List<TextBox> repTextboxes = new List<TextBox>();
        private List<string> currentSubjectsForReps = new List<string>();

        private Button btnCheckUpdate;
        private Button btnSkipVersion;
        private Label lblUpdateContent;
        private Label lblMirrorStatus;
        private ComboBox cmbMirrorManual;

        // 新增：下载进度控件
        private ProgressBar progressDownload;
        private Label lblDownloadStatus;

        private Button btnOK;
        private Button btnCancel;
        private AppConfig config;
        private HomeworkViewer mainForm;

        private int _selectedPage = 0;
        private Version _currentVersion;
        private string _currentVersionString;

        private MirrorManager _mirrorManager;
        private DownloadHelper _downloadHelper;

        public SettingsForm(HomeworkViewer main)
        {
            mainForm = main;
            config = AppConfig.Load();
            _mirrorManager = new MirrorManager();
            _downloadHelper = new DownloadHelper();

            InitializeComponent();
            LoadSettings();
            ShowPage(0);
            this.Opacity = 0.95;
            this.BackColor = Color.Black;
            this.ForeColor = Color.White;
            GetCurrentVersion();
        }

        private void GetCurrentVersion()
        {
            Assembly assembly = Assembly.GetEntryAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string fullVersion = fvi.ProductVersion; // 可能包含 + 和哈希值，如 "1.2.0+4c8b50..."

            // 提取有效的版本号部分（取第一个加号之前的部分）
            string versionPart = fullVersion;
            int plusIndex = fullVersion.IndexOf('+');
            if (plusIndex > 0)
                versionPart = fullVersion.Substring(0, plusIndex);
            int spaceIndex = fullVersion.IndexOf(' ');
            if (spaceIndex > 0)
                versionPart = fullVersion.Substring(0, spaceIndex);

            string[] parts = versionPart.Split('.');
            if (parts.Length >= 3 && 
                int.TryParse(parts[0], out int major) && 
                int.TryParse(parts[1], out int minor) && 
                int.TryParse(parts[2], out int build))
            {
                _currentVersion = new Version(major, minor, build);
                _currentVersionString = $"{major}.{minor}.{build}";
            }
            else
            {
                _currentVersion = new Version(0, 0, 0);
                _currentVersionString = "0.0.0";
            }
        }

        private async Task<RemoteVersionInfo> FetchRemoteVersionInfoAsync()
        {
            string originalUrl = "https://raw.githubusercontent.com/Suixiuliang/HomeworkViewer/refs/heads/main/Version.txt";
            
            // 获取镜像URL
            string url = await _mirrorManager.GetMirroredUrlAsync(originalUrl);
            
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    client.Timeout = TimeSpan.FromSeconds(5);
                    string json = await client.GetStringAsync(url);

                    // 尝试解析为数组
                    try
                    {
                        var parts = JsonSerializer.Deserialize<string[]>(json);
                        if (parts != null && parts.Length == 3)
                        {
                            return new RemoteVersionInfo
                            {
                                Version = parts[0],
                                ReleaseType = parts[1],
                                IsMandatory = parts[2]
                            };
                        }
                    }
                    catch { }

                    // 尝试解析为对象
                    try
                    {
                        var obj = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                        if (obj != null && obj.ContainsKey("Version") && obj.ContainsKey("ReleaseType") && obj.ContainsKey("IsMandatory"))
                        {
                            return new RemoteVersionInfo
                            {
                                Version = obj["Version"],
                                ReleaseType = obj["ReleaseType"],
                                IsMandatory = obj["IsMandatory"]
                            };
                        }
                    }
                    catch { }

                    // 如果都不成功，显示实际内容
                    if (!this.IsDisposed)
                    {
                        this.Invoke((Action)(() =>
                        {
                            MessageBox.Show($"远程版本文件格式不正确。实际内容为：\n{json}\n\n请确保文件是有效的 JSON 数组或对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }));
                    }
                }
                catch (HttpRequestException ex) when (url != originalUrl)
                {
                    // 如果镜像失败，尝试原始URL
                    try
                    {
                        string json = await client.GetStringAsync(originalUrl);
                        // 再次尝试解析
                        try
                        {
                            var parts = JsonSerializer.Deserialize<string[]>(json);
                            if (parts != null && parts.Length == 3)
                            {
                                return new RemoteVersionInfo
                                {
                                    Version = parts[0],
                                    ReleaseType = parts[1],
                                    IsMandatory = parts[2]
                                };
                            }
                        }
                        catch { }
                        try
                        {
                            var obj = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                            if (obj != null && obj.ContainsKey("Version") && obj.ContainsKey("ReleaseType") && obj.ContainsKey("IsMandatory"))
                            {
                                return new RemoteVersionInfo
                                {
                                    Version = obj["Version"],
                                    ReleaseType = obj["ReleaseType"],
                                    IsMandatory = obj["IsMandatory"]
                                };
                            }
                        }
                        catch { }
                        if (!this.IsDisposed)
                        {
                            this.Invoke((Action)(() =>
                            {
                                MessageBox.Show($"远程版本文件格式不正确。实际内容为：\n{json}\n\n请确保文件是有效的 JSON 数组或对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }));
                        }
                    }
                    catch
                    {
                        if (!this.IsDisposed)
                            this.Invoke((Action)(() => MessageBox.Show($"网络错误：{ex.Message}\n请检查网络连接或稍后重试。", "检查更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                    }
                }
                catch (HttpRequestException ex)
                {
                    if (!this.IsDisposed)
                        this.Invoke((Action)(() => MessageBox.Show($"网络错误：{ex.Message}\n请检查网络连接或稍后重试。", "检查更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                }
                catch (TaskCanceledException)
                {
                    if (!this.IsDisposed)
                        this.Invoke((Action)(() => MessageBox.Show("连接超时，请检查网络。", "检查更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                }
                catch (Exception ex)
                {
                    if (!this.IsDisposed)
                        this.Invoke((Action)(() => MessageBox.Show($"未知错误：{ex.Message}", "检查更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                }
            }
            return null;
        }

        private async void btnCheckUpdate_Click(object sender, EventArgs e)
        {
            btnCheckUpdate.Enabled = false;
            btnCheckUpdate.Text = "检查中...";

            RemoteVersionInfo remoteInfo = await FetchRemoteVersionInfoAsync();

            if (this.IsDisposed) return;

            // 保存当前版本引用，避免闭包问题
            Version currentVer = _currentVersion;
            string currentVerStr = _currentVersionString;

            this.Invoke((Action)(() =>
            {
                // 确保控件仍然存在
                if (btnCheckUpdate == null || btnSkipVersion == null || lblUpdateContent == null)
                    return;

                if (remoteInfo != null)
                {
                    Version remoteVersion;
                    if (Version.TryParse(remoteInfo.Version, out remoteVersion) && currentVer != null)
                    {
                        int comparison = currentVer.CompareTo(remoteVersion);
                        if (comparison < 0)
                        {
                            bool isMandatory = remoteInfo.IsMandatory == "0";
                            if (isMandatory)
                            {
                                DialogResult result = MessageBox.Show($"发现强制更新 {remoteInfo.Version} ({remoteInfo.ReleaseType})。\n点击确定前往下载页面。", "强制更新", MessageBoxButtons.OKCancel);
                                if (result == DialogResult.OK)
                                {
                                    Process.Start("https://github.com/Suixiuliang/HomeworkViewer/releases");
                                }
                                btnCheckUpdate.Text = "检查更新";
                                btnCheckUpdate.Enabled = true;
                            }
                            else
                            {
                                btnCheckUpdate.Text = $"更新至 {remoteInfo.Version}";
                                btnCheckUpdate.Click -= btnCheckUpdate_Click;
                                btnCheckUpdate.Click += async (s, ev) =>
                                {
                                    btnCheckUpdate.Enabled = false;
                                    btnCheckUpdate.Text = "准备下载...";

                                    // 显示进度控件
                                    progressDownload.Visible = true;
                                    lblDownloadStatus.Visible = true;
                                    lblDownloadStatus.Text = "准备下载...";
                                    progressDownload.Value = 0;

                                    try
                                    {
                                        var assets = await _downloadHelper.GetLatestReleaseAssetsAsync("Suixiuliang", "HomeworkViewer");
                                        if (assets.Count > 0)
                                        {
                                            var archInfo = _downloadHelper.GetCurrentSystemArch();
                                            var bestAsset = _downloadHelper.FindBestMatchAsset(assets, archInfo);

                                            if (bestAsset != null)
                                            {
                                                var confirm = MessageBox.Show(
                                                    $"找到匹配的安装包：{bestAsset.Name}\n大小：{bestAsset.Size / 1024 / 1024:F2} MB\n\n是否开始下载？",
                                                    "确认下载",
                                                    MessageBoxButtons.YesNo,
                                                    MessageBoxIcon.Question);

                                                if (confirm == DialogResult.Yes)
                                                {
                                                    btnCheckUpdate.Text = "下载中...";

                                                    // 使用进度报告更新UI
                                                    var progress = new Progress<int>(p =>
                                                    {
                                                        if (!this.IsDisposed)
                                                        {
                                                            this.Invoke((Action)(() =>
                                                            {
                                                                progressDownload.Value = Math.Min(p, 100);
                                                                lblDownloadStatus.Text = $"下载中 {p}%";
                                                            }));
                                                        }
                                                    });

                                                    string downloadedFile = await _downloadHelper.DownloadFileAsync(
                                                        bestAsset.DownloadUrl,
                                                        bestAsset.Name,
                                                        progress);

                                                    // 下载完成，隐藏进度控件
                                                    progressDownload.Visible = false;
                                                    lblDownloadStatus.Visible = false;
                                                    btnCheckUpdate.Text = "下载完成";

                                                    _downloadHelper.OpenOrInstallFile(downloadedFile);
                                                }
                                                else
                                                {
                                                    // 用户取消下载，恢复按钮
                                                    btnCheckUpdate.Enabled = true;
                                                    btnCheckUpdate.Text = "检查更新";
                                                    progressDownload.Visible = false;
                                                    lblDownloadStatus.Visible = false;
                                                }
                                            }
                                            else
                                            {
                                                var result = MessageBox.Show(
                                                    "未能自动匹配适合您系统的安装包。是否前往Releases页面手动下载？",
                                                    "提示",
                                                    MessageBoxButtons.YesNo,
                                                    MessageBoxIcon.Information);
                                                if (result == DialogResult.Yes)
                                                    Process.Start("https://github.com/Suixiuliang/HomeworkViewer/releases");
                                                btnCheckUpdate.Enabled = true;
                                                btnCheckUpdate.Text = "检查更新";
                                                progressDownload.Visible = false;
                                                lblDownloadStatus.Visible = false;
                                            }
                                        }
                                        else
                                        {
                                            Process.Start("https://github.com/Suixiuliang/HomeworkViewer/releases");
                                            btnCheckUpdate.Enabled = true;
                                            btnCheckUpdate.Text = "检查更新";
                                            progressDownload.Visible = false;
                                            lblDownloadStatus.Visible = false;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"下载失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        btnCheckUpdate.Enabled = true;
                                        btnCheckUpdate.Text = "检查更新";
                                        progressDownload.Visible = false;
                                        lblDownloadStatus.Visible = false;
                                    }
                                };

                                btnSkipVersion.Visible = true;
                                btnSkipVersion.Text = $"跳过 {remoteInfo.Version}";

                                _ = LoadAndDisplayReleaseNotes(remoteInfo.Version);
                            }
                        }
                        else if (comparison == 0)
                        {
                            MessageBox.Show("当前已是最新版本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            btnCheckUpdate.Text = "检查更新";
                            btnCheckUpdate.Enabled = true;
                        }
                        else
                        {
                            MessageBox.Show($"当前版本 ({currentVerStr}) 高于远程版本 ({remoteInfo.Version})，可能是测试版本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            btnCheckUpdate.Text = "检查更新";
                            btnCheckUpdate.Enabled = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("远程版本号格式无效或当前版本未初始化。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        btnCheckUpdate.Text = "检查更新";
                        btnCheckUpdate.Enabled = true;
                    }
                }
                else
                {
                    btnCheckUpdate.Text = "检查更新";
                    btnCheckUpdate.Enabled = true;
                }
            }));
        }

        private async Task LoadAndDisplayReleaseNotes(string targetVersion)
        {
            if (this.IsDisposed) return;
            this.Invoke((Action)(() => {
                if (lblUpdateContent != null)
                    lblUpdateContent.Text = "加载更新内容中...";
            }));
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string apiUrl = "https://api.github.com/repos/Suixiuliang/HomeworkViewer/releases";
                    client.DefaultRequestHeaders.Add("User-Agent", "HomeworkViewer-Updater");
                    string json = await client.GetStringAsync(apiUrl);
                    if (this.IsDisposed) return;
                    this.Invoke((Action)(() => {
                        if (lblUpdateContent != null)
                            lblUpdateContent.Text = "请前往 Releases 页面查看详细更新内容。";
                    }));
                }
                catch
                {
                    if (this.IsDisposed) return;
                    this.Invoke((Action)(() => {
                        if (lblUpdateContent != null)
                            lblUpdateContent.Text = "无法加载更新内容。";
                    }));
                }
            }
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
            
            btnAbout = CreateSidebarButton("关于", 4);
            btnReps = CreateSidebarButton("科代表", 3);
            btnTimeSlot = CreateSidebarButton("时间段", 2);
            btnAppearance = CreateSidebarButton("外观", 1);
            btnBasic = CreateSidebarButton("基本设置", 0);

            sidebar.Controls.Add(btnAbout);
            sidebar.Controls.Add(btnReps);
            sidebar.Controls.Add(btnTimeSlot);
            sidebar.Controls.Add(btnAppearance);
            sidebar.Controls.Add(btnBasic);

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

            Label lblFontColor = new Label { Text = "字体颜色:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblFontColor, 0, 2);
            FlowLayoutPanel colorPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Height = 30, Width = 200, BackColor = Color.Transparent };
            rbBlack = new RadioButton { Text = "黑色", Checked = true, AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            rbWhite = new RadioButton { Text = "白色", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            colorPanel.Controls.Add(rbBlack);
            colorPanel.Controls.Add(rbWhite);
            layout.Controls.Add(colorPanel, 1, 2);

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
        private void BtnTestFlash_Click(object sender, EventArgs e) => mainForm.StartDebugFlashing();
        private void BtnStopFlash_Click(object sender, EventArgs e) => mainForm.StopDebugFlashing();
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
                RowCount = 8, // 增加两行用于进度控件
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

            // 镜像状态显示
            lblMirrorStatus = new Label
            {
                Text = "正在检测镜像站...",
                AutoSize = true,
                ForeColor = Color.LightGreen,
                BackColor = Color.Transparent,
                Font = new Font("微软雅黑", 9)
            };
            layout.Controls.Add(lblMirrorStatus, 0, 3);

            // 手动选择镜像下拉框
            cmbMirrorManual = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 200,
                Height = 25,
                Visible = true,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cmbMirrorManual.SelectedIndexChanged += async (s, e) =>
            {
                if (cmbMirrorManual.SelectedItem != null && cmbMirrorManual.SelectedIndex > 0)
                {
                    _mirrorManager.ClearCache();
                    await UpdateMirrorStatusAsync(lblMirrorStatus);
                }
            };
            layout.Controls.Add(cmbMirrorManual, 0, 4);

            // 下载进度条和状态标签
            progressDownload = new ProgressBar
            {
                Width = 200,
                Height = 20,
                Visible = false,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100
            };
            layout.Controls.Add(progressDownload, 0, 5);

            lblDownloadStatus = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Font = new Font("微软雅黑", 9),
                Visible = false
            };
            layout.Controls.Add(lblDownloadStatus, 0, 6);

            // 检查更新按钮和跳过按钮
            FlowLayoutPanel updatePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent, Margin = new Padding(0, 10, 0, 0) };
            btnCheckUpdate = new Button { Text = "检查更新", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnCheckUpdate.Click += btnCheckUpdate_Click;

            btnSkipVersion = new Button { Text = "跳过", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Visible = false };
            btnSkipVersion.Click += (s, e) => { btnSkipVersion.Visible = false; btnCheckUpdate.Text = "检查更新"; };

            lblUpdateContent = new Label { Text = "", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent, Font = new Font("微软雅黑", 9) };

            updatePanel.Controls.Add(btnCheckUpdate);
            updatePanel.Controls.Add(btnSkipVersion);
            layout.Controls.Add(updatePanel, 0, 7);
            layout.Controls.Add(lblUpdateContent, 0, 8);

            aboutPanel.Controls.Add(layout);

            // 异步更新镜像状态
            _ = UpdateMirrorStatusAsync(lblMirrorStatus);
        }

        private async Task UpdateMirrorStatusAsync(Label statusLabel)
        {
            if (this.IsDisposed) return;

            try
            {
                var mirrors = await _mirrorManager.GetWorkingMirrorsAsync();
                string status = mirrors.Any()
                    ? $"✅ 可用镜像站: {mirrors.Count}个，最快: {mirrors.First()}"
                    : "⚠️ 未检测到可用镜像站，将使用原始GitHub";

                this.Invoke((Action)(() =>
                {
                    statusLabel.Text = status;

                    cmbMirrorManual.Items.Clear();
                    cmbMirrorManual.Items.Add("自动选择（推荐）");
                    foreach (var mirror in mirrors)
                    {
                        cmbMirrorManual.Items.Add(mirror);
                    }
                    cmbMirrorManual.SelectedIndex = 0;
                }));
            }
            catch
            {
                this.Invoke((Action)(() =>
                {
                    statusLabel.Text = "⚠️ 镜像检测失败";
                }));
            }
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