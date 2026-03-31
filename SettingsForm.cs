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
using System.IO;

namespace HomeworkViewer
{
    public class RemoteVersionInfo
    {
        public string Version { get; set; }
        public string ReleaseType { get; set; }
        public string IsMandatory { get; set; }
        public string Hash { get; set; }
        public string DownloadUrl { get; set; }
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

        private ComboBox cmbBgEffect;

        private NumericUpDown numEveningCount;
        private FlowLayoutPanel eveningPanel;
        private List<DateTimePicker> startPickers = new List<DateTimePicker>();
        private List<DateTimePicker> endPickers = new List<DateTimePicker>();
        private Button btnTestFlash;
        private Button btnStopFlash;

        private FlowLayoutPanel repsPanel;
        private List<TextBox> repTextboxes = new List<TextBox>();
        private List<string> currentSubjectsForReps = new List<string>();

        // 字体相关
        private ComboBox cmbFontFamily;
        private Button btnSelectCustomFont;

        // 导出相关
        private ComboBox cmbDefaultExportFormat;

        private Button btnCheckUpdate;
        private Button btnSkipVersion;
        private Label lblUpdateContent;
        private Label lblMirrorStatus;
        private ComboBox cmbMirrorManual;
        private ProgressBar progressDownload;
        private Label lblDownloadStatus;

        private Button btnOK;
        private Button btnCancel;
        private AppConfig config;
        private HomeworkViewer mainForm;
        private CheckBox chkEnableMarkdown;

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

            this.Load += (s, e) => { _ = UpdateMirrorStatusAsync(lblMirrorStatus); };
        }

        private void GetCurrentVersion()
        {
            Assembly assembly = Assembly.GetEntryAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string fullVersion = fvi.ProductVersion;

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
            string url = await _mirrorManager.GetMirroredUrlAsync(originalUrl);

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    client.Timeout = TimeSpan.FromSeconds(5);
                    string json = await client.GetStringAsync(url);

                    try
                    {
                        var parts = JsonSerializer.Deserialize<string[]>(json);
                        if (parts != null && parts.Length >= 3)
                        {
                            var info = new RemoteVersionInfo
                            {
                                Version = parts[0],
                                ReleaseType = parts[1],
                                IsMandatory = parts[2]
                            };
                            if (parts.Length >= 4)
                                info.Hash = parts[3];
                            if (parts.Length >= 5)
                                info.DownloadUrl = parts[4];
                            return info;
                        }
                    }
                    catch { }

                    try
                    {
                        var obj = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                        if (obj != null && obj.ContainsKey("Version") && obj.ContainsKey("ReleaseType") && obj.ContainsKey("IsMandatory"))
                        {
                            var info = new RemoteVersionInfo
                            {
                                Version = obj["Version"],
                                ReleaseType = obj["ReleaseType"],
                                IsMandatory = obj["IsMandatory"]
                            };
                            if (obj.ContainsKey("Hash"))
                                info.Hash = obj["Hash"];
                            if (obj.ContainsKey("DownloadUrl"))
                                info.DownloadUrl = obj["DownloadUrl"];
                            return info;
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
                catch (HttpRequestException _) when (url != originalUrl)
                {
                    try
                    {
                        string json = await client.GetStringAsync(originalUrl);
                        try
                        {
                            var parts = JsonSerializer.Deserialize<string[]>(json);
                            if (parts != null && parts.Length >= 3)
                            {
                                var info = new RemoteVersionInfo
                                {
                                    Version = parts[0],
                                    ReleaseType = parts[1],
                                    IsMandatory = parts[2]
                                };
                                if (parts.Length >= 4)
                                    info.Hash = parts[3];
                                if (parts.Length >= 5)
                                    info.DownloadUrl = parts[4];
                                return info;
                            }
                        }
                        catch { }
                        try
                        {
                            var obj = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                            if (obj != null && obj.ContainsKey("Version") && obj.ContainsKey("ReleaseType") && obj.ContainsKey("IsMandatory"))
                            {
                                var info = new RemoteVersionInfo
                                {
                                    Version = obj["Version"],
                                    ReleaseType = obj["ReleaseType"],
                                    IsMandatory = obj["IsMandatory"]
                                };
                                if (obj.ContainsKey("Hash"))
                                    info.Hash = obj["Hash"];
                                if (obj.ContainsKey("DownloadUrl"))
                                    info.DownloadUrl = obj["DownloadUrl"];
                                return info;
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
                            this.Invoke((Action)(() => MessageBox.Show($"网络错误，请检查网络连接或稍后重试。", "检查更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error)));
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

            Version currentVer = _currentVersion;
            string currentVerStr = _currentVersionString;

            this.Invoke((Action)(() =>
            {
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
                                if (!string.IsNullOrEmpty(remoteInfo.DownloadUrl))
                                {
                                    btnCheckUpdate.Text = "准备下载...";
                                    progressDownload.Visible = true;
                                    lblDownloadStatus.Visible = true;
                                    lblDownloadStatus.Text = "准备下载...";
                                    progressDownload.Value = 0;

                                    Task.Run(async () =>
                                    {
                                        try
                                        {
                                            string fileName = remoteInfo.DownloadUrl.Substring(remoteInfo.DownloadUrl.LastIndexOf('/') + 1);
                                            if (string.IsNullOrEmpty(fileName))
                                                fileName = "update.exe";

                                            var confirm = MessageBox.Show(
                                                $"发现强制更新 {remoteInfo.Version} ({remoteInfo.ReleaseType})。\n将下载更新包：{fileName}\n\n是否立即下载并更新？",
                                                "强制更新",
                                                MessageBoxButtons.YesNo,
                                                MessageBoxIcon.Warning);

                                            if (confirm == DialogResult.Yes)
                                            {
                                                this.Invoke((Action)(() => btnCheckUpdate.Text = "下载中..."));

                                                var progress = new Progress<DownloadProgressInfo>(p =>
                                                {
                                                    if (!this.IsDisposed)
                                                    {
                                                        this.Invoke((Action)(() =>
                                                        {
                                                            progressDownload.Value = Math.Min(p.Percentage, 100);
                                                            string speed = p.SpeedKBps > 0 ? $"{p.SpeedKBps:F1} KB/s" : "计算中...";
                                                            string remaining = p.TimeRemaining > TimeSpan.Zero ? $"{p.TimeRemaining:mm\\:ss}" : "--:--";
                                                            lblDownloadStatus.Text = $"下载中 {p.Percentage}% | {speed} | 剩余 {remaining}";
                                                        }));
                                                    }
                                                });

                                                string expectedHash = null;
                                                if (remoteInfo.Hash != null && remoteInfo.Hash.StartsWith("sha256:"))
                                                    expectedHash = remoteInfo.Hash.Substring(7);

                                                string downloadedFile = await _downloadHelper.DownloadFileAsync(
                                                    remoteInfo.DownloadUrl,
                                                    fileName,
                                                    expectedHash,
                                                    progress);

                                                this.Invoke((Action)(() =>
                                                {
                                                    progressDownload.Visible = false;
                                                    lblDownloadStatus.Visible = false;
                                                    btnCheckUpdate.Text = "下载完成";
                                                }));

                                                config.UpdatePending = 1;
                                                config.Save();

                                                DialogResult dialogResult = MessageBox.Show(
                                                    "更新包已下载完成。请点击确定启动安装程序，应用将自动关闭。",
                                                    "更新",
                                                    MessageBoxButtons.OKCancel,
                                                    MessageBoxIcon.Information);

                                                if (dialogResult == DialogResult.OK)
                                                {
                                                    _downloadHelper.OpenOrInstallFile(downloadedFile);
                                                    Application.Exit();
                                                }
                                                else
                                                {
                                                    this.Invoke((Action)(() =>
                                                    {
                                                        btnCheckUpdate.Enabled = true;
                                                        btnCheckUpdate.Text = "检查更新";
                                                    }));
                                                }
                                            }
                                            else
                                            {
                                                this.Invoke((Action)(() =>
                                                {
                                                    btnCheckUpdate.Enabled = true;
                                                    btnCheckUpdate.Text = "检查更新";
                                                    progressDownload.Visible = false;
                                                    lblDownloadStatus.Visible = false;
                                                }));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"下载失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            this.Invoke((Action)(() =>
                                            {
                                                btnCheckUpdate.Enabled = true;
                                                btnCheckUpdate.Text = "检查更新";
                                                progressDownload.Visible = false;
                                                lblDownloadStatus.Visible = false;
                                            }));
                                        }
                                    });
                                }
                                else
                                {
                                    DialogResult result = MessageBox.Show($"发现强制更新 {remoteInfo.Version} ({remoteInfo.ReleaseType})。\n点击确定前往下载页面。", "强制更新", MessageBoxButtons.OKCancel);
                                    if (result == DialogResult.OK)
                                    {
                                        Process.Start(new ProcessStartInfo
                                        {
                                            FileName = "https://github.com/Suixiuliang/HomeworkViewer/releases",
                                            UseShellExecute = true
                                        });
                                    }
                                    btnCheckUpdate.Text = "检查更新";
                                    btnCheckUpdate.Enabled = true;
                                }
                            }
                            else
                            {
                                btnCheckUpdate.Text = $"更新至 {remoteInfo.Version}";
                                btnCheckUpdate.Click -= btnCheckUpdate_Click;
                                btnCheckUpdate.Click += async (s, ev) =>
                                {
                                    btnCheckUpdate.Enabled = false;
                                    btnCheckUpdate.Text = "准备下载...";

                                    progressDownload.Visible = true;
                                    lblDownloadStatus.Visible = true;
                                    lblDownloadStatus.Text = "准备下载...";
                                    progressDownload.Value = 0;

                                    try
                                    {
                                        if (!string.IsNullOrEmpty(remoteInfo.DownloadUrl))
                                        {
                                            string fileName = remoteInfo.DownloadUrl.Substring(remoteInfo.DownloadUrl.LastIndexOf('/') + 1);
                                            if (string.IsNullOrEmpty(fileName))
                                                fileName = "update.exe";

                                            var confirm = MessageBox.Show(
                                                $"将下载更新包：{fileName}\n\n是否开始下载？",
                                                "确认下载",
                                                MessageBoxButtons.YesNo,
                                                MessageBoxIcon.Question);

                                            if (confirm == DialogResult.Yes)
                                            {
                                                btnCheckUpdate.Text = "下载中...";

                                                var progress = new Progress<DownloadProgressInfo>(p =>
                                                {
                                                    if (!this.IsDisposed)
                                                    {
                                                        this.Invoke((Action)(() =>
                                                        {
                                                            progressDownload.Value = Math.Min(p.Percentage, 100);
                                                            string speed = p.SpeedKBps > 0 ? $"{p.SpeedKBps:F1} KB/s" : "计算中...";
                                                            string remaining = p.TimeRemaining > TimeSpan.Zero ? $"{p.TimeRemaining:mm\\:ss}" : "--:--";
                                                            lblDownloadStatus.Text = $"下载中 {p.Percentage}% | {speed} | 剩余 {remaining}";
                                                        }));
                                                    }
                                                });

                                                string expectedHash = null;
                                                if (remoteInfo.Hash != null && remoteInfo.Hash.StartsWith("sha256:"))
                                                    expectedHash = remoteInfo.Hash.Substring(7);

                                                string downloadedFile = await _downloadHelper.DownloadFileAsync(
                                                    remoteInfo.DownloadUrl,
                                                    fileName,
                                                    expectedHash,
                                                    progress);

                                                progressDownload.Visible = false;
                                                lblDownloadStatus.Visible = false;
                                                btnCheckUpdate.Text = "下载完成";

                                                config.UpdatePending = 1;
                                                config.Save();

                                                DialogResult dialogResult = MessageBox.Show(
                                                    "更新包已下载完成。请点击确定启动安装程序，应用将自动关闭。",
                                                    "更新",
                                                    MessageBoxButtons.OKCancel,
                                                    MessageBoxIcon.Information);

                                                if (dialogResult == DialogResult.OK)
                                                {
                                                    _downloadHelper.OpenOrInstallFile(downloadedFile);
                                                    Application.Exit();
                                                }
                                                else
                                                {
                                                    btnCheckUpdate.Enabled = true;
                                                    btnCheckUpdate.Text = "检查更新";
                                                }
                                            }
                                            else
                                            {
                                                btnCheckUpdate.Enabled = true;
                                                btnCheckUpdate.Text = "检查更新";
                                                progressDownload.Visible = false;
                                                lblDownloadStatus.Visible = false;
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("更新信息中缺少下载链接，请联系开发者。", "下载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                btnCheckUpdate.Enabled = true;
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
            this.Size = new Size(650, 700);
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
                RowCount = 5,
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

            Label lblExportFormat = new Label { Text = "默认导出格式:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblExportFormat, 0, 3);
            cmbDefaultExportFormat = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 150,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cmbDefaultExportFormat.Items.AddRange(new object[] { "txt", "pdf", "jpg", "html" });
            cmbDefaultExportFormat.SelectedItem = config.ExportFormat;
            layout.Controls.Add(cmbDefaultExportFormat, 1, 3);

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
                RowCount = 8,
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

            // 顶部条颜色
            Label lblBarColor = new Label { Text = "顶部条颜色:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblBarColor, 0, 3);
            FlowLayoutPanel barColorPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Height = 30, Width = 250, BackColor = Color.Transparent };
            btnBarColor = new Button { Text = "选择颜色", Width = 80, Height = 25, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnBarColor.Click += BtnBarColor_Click;
            pnlBarColorPreview = new Panel { Width = 40, Height = 25, BackColor = Color.Yellow, BorderStyle = BorderStyle.FixedSingle };
            barColorPanel.Controls.Add(btnBarColor);
            barColorPanel.Controls.Add(pnlBarColorPreview);
            layout.Controls.Add(barColorPanel, 1, 3);

            // 背景效果
            Label lblBgEffect = new Label { Text = "背景效果:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblBgEffect, 0, 4);
            cmbBgEffect = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 150,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cmbBgEffect.Items.AddRange(new object[] { "Mica", "Acrylic", "Aero" });
            cmbBgEffect.SelectedItem = config.BackgroundEffect;
            layout.Controls.Add(cmbBgEffect, 1, 4);

            // Markdown 渲染
            Label lblMarkdown = new Label { Text = "Markdown 渲染:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblMarkdown, 0, 5);
            chkEnableMarkdown = new CheckBox
            {
                Text = "启用 Markdown 渲染（不稳定）",
                Checked = config.EnableMarkdown,
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };
            layout.Controls.Add(chkEnableMarkdown, 1, 5);

            // 字体选择
            Label lblFontFamily = new Label { Text = "字体:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, BackColor = Color.Transparent, AutoSize = true };
            layout.Controls.Add(lblFontFamily, 0, 6);
            FlowLayoutPanel fontPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Height = 30, Width = 250, BackColor = Color.Transparent };
            cmbFontFamily = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 150,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cmbFontFamily.Items.AddRange(new object[] { "微软雅黑", "黑体", "楷体", "苹方", "自定义" });
            cmbFontFamily.SelectedItem = config.IsCustomFont ? "自定义" : config.FontFamily;
            btnSelectCustomFont = new Button
            {
                Text = "选择字体文件",
                Width = 100,
                Height = 25,
                Visible = false,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnSelectCustomFont.Click += BtnSelectCustomFont_Click;
            cmbFontFamily.SelectedIndexChanged += (s, e) =>
            {
                btnSelectCustomFont.Visible = cmbFontFamily.SelectedItem?.ToString() == "自定义";
            };
            fontPanel.Controls.Add(cmbFontFamily);
            fontPanel.Controls.Add(btnSelectCustomFont);
            layout.Controls.Add(fontPanel, 1, 6);

            appearancePanel.Controls.Add(layout);
        }

        private void BtnSelectCustomFont_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "字体文件|*.ttf;*.ttc";
                ofd.Title = "选择自定义字体文件";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string destPath = Path.Combine(Application.StartupPath, "selfdeffont" + Path.GetExtension(ofd.FileName));
                    try
                    {
                        File.Copy(ofd.FileName, destPath, true);
                        config.FontFamily = destPath;
                        config.IsCustomFont = true;
                        FontManager.LoadCustomFont(destPath);
                        MessageBox.Show("自定义字体已加载，重启应用后生效。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"字体复制失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
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
            btnTestFlash = new Button { Text = "测试闪烁", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Visible = false };
            btnTestFlash.Click += BtnTestFlash_Click;
            btnStopFlash = new Button { Text = "停止闪烁", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0), Visible = false };
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
                RowCount = 8,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            Label lblVersion = new Label { Text = "作业展板 版本 1.5.5", Font = new Font("微软雅黑", 12, FontStyle.Bold), AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            layout.Controls.Add(lblVersion, 0, 0);
            Label lblAuthor = new Label { Text = "\n作者: MaxSui 隋修梁", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent, Font = new Font("微软雅黑", 10) };
            layout.Controls.Add(lblAuthor, 0, 1);
            Label lblCopyright = new Label { Text = "\n© 2026 MaxSui 保留部分权利 \n本软件遵循GNU General Public License 3.0开源协议", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            layout.Controls.Add(lblCopyright, 0, 2);

            lblMirrorStatus = new Label
            {
                Text = "正在检测镜像站...",
                AutoSize = true,
                ForeColor = Color.LightGreen,
                BackColor = Color.Transparent,
                Font = new Font("微软雅黑", 9)
            };
            layout.Controls.Add(lblMirrorStatus, 0, 3);

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
            cmbDefaultExportFormat.SelectedItem = config.ExportFormat;
            cmbFontFamily.SelectedItem = config.IsCustomFont ? "自定义" : config.FontFamily;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            config.LastMode = cmbMode.SelectedItem?.ToString() ?? "大理";
            config.FontSizeLevel = cmbFontSize.SelectedIndex;
            config.CardOpacity = trackCardOpacity.Value;
            config.BackgroundOpacity = trackBgOpacity.Value;
            config.FontColorWhite = rbWhite.Checked;
            config.ScrollSpeed = (int)numScrollSpeed.Value;
            config.BackgroundEffect = cmbBgEffect.SelectedItem?.ToString() ?? "Mica";
            config.EnableMarkdown = chkEnableMarkdown.Checked;
            config.ExportFormat = cmbDefaultExportFormat.SelectedItem?.ToString() ?? "txt";

            string fontSelection = cmbFontFamily.SelectedItem?.ToString();
            if (fontSelection == "自定义")
            {
                config.IsCustomFont = true;
            }
            else
            {
                config.IsCustomFont = false;
                config.FontFamily = fontSelection;
            }

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

            // 移除 Markdown 启用时的对话框
            mainForm.ApplySettings(config);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}