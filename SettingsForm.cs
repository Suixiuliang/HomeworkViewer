#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
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
        // 自定义控件：绘制 1 像素高的细线，支持透明背景
        private class LineControl : Control
        {
            public LineControl()
            {
                SetStyle(ControlStyles.SupportsTransparentBackColor, true);
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.ResizeRedraw, true);
                Height = 1;
                BackColor = Color.Transparent;
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using (var pen = new Pen(Color.FromArgb(100, 100, 100), 1))
                {
                    e.Graphics.DrawLine(pen, 0, 0, Width, 0);
                }
            }
        }

        private Button btnBasic;
        private Button btnAppearance;
        private Button btnTimeSlot;
        private Button btnMgmt;
        private Button btnAbout;
        private Panel contentPanel;
        private Panel sidebar;

        // 滑块动画相关
        private RectangleF _selectorRect;
        private float _selectorTargetY;
        private float _selectorStartY;
        private float _selectorCurrentY;
        private float _selectorAnimT;
        private System.Windows.Forms.Timer _selectorTimer;
        private const int SELECTOR_CORNER_RADIUS = 9;
        private readonly Color SELECTOR_COLOR = Color.FromArgb(63, 191, 255);
        private List<Button> _sidebarButtons = new List<Button>();

        // 基本设置
        private ComboBox cmbMode;
        private ComboBox cmbFontSize;
        private NumericUpDown numScrollSpeed;
        private ComboBox cmbDefaultExportFormat;

        // 外观
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
        private ComboBox cmbFontFamily;
        private Button btnSelectCustomFont;
        private CheckBox chkShowMouseGlow;

        // 背景图片相关
        private RadioButton rbTransparentBg;
        private RadioButton rbImageBg;
        private TextBox txtBgImagePath;
        private Button btnBrowseBgImage;

        // 时间段
        private NumericUpDown numEveningCount;
        private FlowLayoutPanel eveningPanel;
        private List<DateTimePicker> startPickers = new List<DateTimePicker>();
        private List<DateTimePicker> endPickers = new List<DateTimePicker>();
        private Button btnTestFlash;
        private Button btnStopFlash;
        private CheckBox chkShowDueTime;

        // 集控管理
        private CheckBox chkMgmtEnabled;
        private TextBox txtMgmtManifestUrl;
        private Button btnTestMgmtConnection;
        private Label lblMgmtStatus;
        private Label lblMgmtOrgName;

        // 关于
        private Button btnCheckUpdate;
        private Button btnSkipVersion;
        private Label lblUpdateContent;
        private Label lblMirrorStatus;
        private ComboBox cmbMirrorManual;
        private ProgressBar progressDownload;
        private Label lblDownloadStatus;
        private Button btnResetAll;
        private CheckBox chkUseWebView2;    // 新增：是否启用 WebView2 渲染

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
            InitSelectorAnimation();
            ShowPage(0);
            this.Opacity = 0.95;
            this.BackColor = Color.Black;
            this.ForeColor = Color.White;
            GetCurrentVersion();

            this.Load += (s, e) => { _ = UpdateMirrorStatusAsync(lblMirrorStatus); };
        }

        private void InitSelectorAnimation()
        {
            _sidebarButtons.Clear();
            _sidebarButtons.Add(btnBasic);
            _sidebarButtons.Add(btnAppearance);
            _sidebarButtons.Add(btnTimeSlot);
            _sidebarButtons.Add(btnMgmt);
            _sidebarButtons.Add(btnAbout);

            if (_sidebarButtons.Count > 0)
            {
                var activeBtn = _sidebarButtons[_selectedPage];
                _selectorRect = new RectangleF(activeBtn.Left, activeBtn.Top, activeBtn.Width, activeBtn.Height);
                _selectorCurrentY = _selectorRect.Y;
                foreach (var btn in _sidebarButtons)
                    btn.ForeColor = Color.White;
                activeBtn.ForeColor = SELECTOR_COLOR;
            }

            _selectorTimer = new System.Windows.Forms.Timer { Interval = 10, Enabled = false };
            _selectorTimer.Tick += SelectorTimer_Tick;
            sidebar.Paint += Sidebar_Paint;
        }

        private void Sidebar_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = GetRoundedRectPath(_selectorRect, SELECTOR_CORNER_RADIUS))
            using (var brush = new SolidBrush(SELECTOR_COLOR))
            {
                e.Graphics.FillPath(brush, path);
            }
        }

        private GraphicsPath GetRoundedRectPath(RectangleF rect, int radius)
        {
            var path = new GraphicsPath();
            float x = rect.X, y = rect.Y, w = rect.Width, h = rect.Height, r = radius, d = r * 2;
            path.AddArc(x, y, d, d, 180, 90);
            path.AddLine(x + r, y, x + w - r, y);
            path.AddArc(x + w - d, y, d, d, 270, 90);
            path.AddLine(x + w, y + r, x + w, y + h - r);
            path.AddArc(x + w - d, y + h - d, d, d, 0, 90);
            path.AddLine(x + w - r, y + h, x + r, y + h);
            path.AddArc(x, y + h - d, d, d, 90, 90);
            path.AddLine(x, y + h - r, x, y + r);
            path.CloseFigure();
            return path;
        }

        private void StartSelectorAnimation(Button targetBtn)
        {
            _selectorStartY = _selectorCurrentY;
            _selectorTargetY = targetBtn.Top;
            _selectorAnimT = 0;
            _selectorTimer.Enabled = true;
        }

        private void SelectorTimer_Tick(object sender, EventArgs e)
        {
            _selectorAnimT += 0.10f;
            float t = Math.Min(1f, _selectorAnimT);
            float eased = (2 - t) * t;
            float newY = _selectorStartY + (_selectorTargetY - _selectorStartY) * eased;
            _selectorCurrentY = newY;
            _selectorRect = new RectangleF(_selectorRect.X, _selectorCurrentY, _selectorRect.Width, _selectorRect.Height);
            sidebar.Invalidate();

            if (t >= 1f)
            {
                _selectorTimer.Enabled = false;
                _selectorRect = new RectangleF(_selectorRect.X, _selectorTargetY, _selectorRect.Width, _selectorRect.Height);
                sidebar.Invalidate();
            }
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
                catch (HttpRequestException) when (url != originalUrl)
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
            this.Size = new Size(650, 780);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            sidebar = new Panel { Dock = DockStyle.Left, Width = 120, BackColor = Color.FromArgb(30, 30, 30) };

            btnAbout = CreateSidebarButton("关于", 4);
            btnMgmt = CreateSidebarButton("集控管理", 3);
            btnTimeSlot = CreateSidebarButton("时间段", 2);
            btnAppearance = CreateSidebarButton("外观", 1);
            btnBasic = CreateSidebarButton("基本设置", 0);

            sidebar.Controls.Add(btnAbout);
            sidebar.Controls.Add(btnMgmt);
            sidebar.Controls.Add(btnTimeSlot);
            sidebar.Controls.Add(btnAppearance);
            sidebar.Controls.Add(btnBasic);

            contentPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(45, 45, 48), Padding = new Padding(20, 20, 20, 20), AutoScroll = true };

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
            CreateMgmtPage();
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
                foreach (var button in _sidebarButtons)
                    button.ForeColor = Color.White;
                btn.ForeColor = SELECTOR_COLOR;
                StartSelectorAnimation(btn);
                ShowPage(pageIndex);
                _selectedPage = pageIndex;
            }
        }

        private void ShowPage(int pageIndex)
        {
            contentPanel.Controls.Clear();
            switch (pageIndex)
            {
                case 0: ShowBasicPage(); break;
                case 1: ShowAppearancePage(); break;
                case 2: ShowTimeSlotPage(); break;
                case 3: ShowMgmtPage(); break;
                case 4: ShowAboutPage(); break;
            }
        }

        // ---------- 基本设置 ----------
        private Panel basicPanel;
        private void CreateBasicPage()
        {
            basicPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent, AutoSize = true };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 4,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 4; i++)
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            AddLabelAndControl(layout, "展示模式:", cmbMode = CreateComboBox(150, new[] { "大理", "中理", "小理", "大文", "全科" }), 0);
            AddLabelAndControl(layout, "字号大小:", cmbFontSize = CreateComboBox(100, new[] { "小", "中", "大" }), 1);
            AddLabelAndControl(layout, "滚动速度(px/s):", numScrollSpeed = new NumericUpDown { Minimum = 0, Maximum = 200, Value = config.ScrollSpeed, Width = 80, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White }, 2);
            AddLabelAndControl(layout, "默认导出格式:", cmbDefaultExportFormat = CreateComboBox(150, new[] { "txt", "pdf", "jpg", "html" }), 3);
            cmbDefaultExportFormat.SelectedItem = config.ExportFormat;

            basicPanel.Controls.Add(layout);
        }
        private void ShowBasicPage() => contentPanel.Controls.Add(basicPanel);

        // ---------- 外观设置 ----------
        private Panel appearancePanel;
        private void CreateAppearancePage()
        {
            appearancePanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent, AutoSize = true };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            int row = 0;

            // 卡片样式
            AddSectionTitle(layout, "卡片样式", ref row);
            trackCardOpacity = new TrackBar { Minimum = 0, Maximum = 100, TickFrequency = 10, Width = 150, BackColor = Color.FromArgb(45, 45, 48) };
            numCardOpacity = new NumericUpDown { Minimum = 0, Maximum = 100, Width = 60, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
            trackCardOpacity.ValueChanged += (s, e) => numCardOpacity.Value = trackCardOpacity.Value;
            numCardOpacity.ValueChanged += (s, e) => trackCardOpacity.Value = (int)numCardOpacity.Value;
            Panel cardPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            cardPanel.Controls.Add(trackCardOpacity);
            cardPanel.Controls.Add(numCardOpacity);
            AddLabelAndControl(layout, "卡片透明度:", cardPanel, row++);
            AddLabelAndControl(layout, "背景透明度:", CreateBgOpacityControl(), row++);
            AddLabelAndControl(layout, "顶部条颜色:", CreateColorPicker(), row++);
            AddSeparator(layout, ref row);

            // 字体设置
            AddSectionTitle(layout, "字体设置", ref row);
            AddLabelAndControl(layout, "字体颜色:", CreateFontColorRadio(), row++);
            AddLabelAndControl(layout, "字体:", CreateFontFamilySelector(), row++);
            AddSeparator(layout, ref row);

            // 视觉效果
            AddSectionTitle(layout, "视觉效果", ref row);
            AddLabelAndControl(layout, "背景效果:", cmbBgEffect = CreateComboBox(150, new[] { "Mica", "Acrylic", "Aero" }), row++);
            cmbBgEffect.SelectedItem = config.BackgroundEffect;
            AddLabelAndControl(layout, "鼠标光晕:", chkShowMouseGlow = new CheckBox { Text = "显示鼠标跟随光晕", Checked = config.ShowMouseGlow, AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent }, row++);
            AddSeparator(layout, ref row);

            // 背景图片
            AddSectionTitle(layout, "背景图片", ref row);
            AddLabelAndControl(layout, "背景类型:", CreateBgTypeRadio(), row++);
            AddLabelAndControl(layout, "图片路径:", CreateImagePathSelector(), row++);

            appearancePanel.Controls.Add(layout);
        }

        private Control CreateBgOpacityControl()
        {
            trackBgOpacity = new TrackBar { Minimum = 0, Maximum = 100, TickFrequency = 10, Width = 150, BackColor = Color.FromArgb(45, 45, 48) };
            numBgOpacity = new NumericUpDown { Minimum = 0, Maximum = 100, Width = 60, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
            trackBgOpacity.ValueChanged += (s, e) => numBgOpacity.Value = trackBgOpacity.Value;
            numBgOpacity.ValueChanged += (s, e) => trackBgOpacity.Value = (int)numBgOpacity.Value;
            FlowLayoutPanel panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            panel.Controls.Add(trackBgOpacity);
            panel.Controls.Add(numBgOpacity);
            return panel;
        }

        private void ShowAppearancePage() => contentPanel.Controls.Add(appearancePanel);

        // ---------- 时间段设置 ----------
        private Panel timeSlotPanel;
        private void CreateTimeSlotPage()
        {
            timeSlotPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent, AutoScroll = true };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 1,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            int row = 0;

            AddSectionTitle(layout, "通用设置", ref row, true);
            layout.Controls.Add(chkShowDueTime = new CheckBox
            {
                Text = "显示提交时间（不勾选则隐藏所有界面的提交时间和提醒）",
                Checked = config.ShowDueTime,
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 5, 0, 10)
            }, 0, row++);

            AddSectionTitle(layout, "晚修节数", ref row, true);
            TableLayoutPanel countLayout = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, AutoSize = true, BackColor = Color.Transparent };
            countLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            countLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            countLayout.Controls.Add(new Label { Text = "节数:", TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.White, AutoSize = true }, 0, 0);
            numEveningCount = new NumericUpDown { Minimum = 1, Maximum = 6, Value = config.EveningClassCount, Width = 80, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
            numEveningCount.ValueChanged += NumEveningCount_ValueChanged;
            countLayout.Controls.Add(numEveningCount, 1, 0);
            layout.Controls.Add(countLayout, 0, row++);

            AddSectionTitle(layout, "各节时间段", ref row, true);
            eveningPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 5, 0, 10)
            };
            layout.Controls.Add(eveningPanel, 0, row++);

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent, Margin = new Padding(0, 10, 0, 0) };
            btnTestFlash = new Button { Text = "测试闪烁", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Visible = false };
            btnTestFlash.Click += BtnTestFlash_Click;
            btnStopFlash = new Button { Text = "停止闪烁", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0), Visible = false };
            btnStopFlash.Click += BtnStopFlash_Click;
            buttonPanel.Controls.Add(btnTestFlash);
            buttonPanel.Controls.Add(btnStopFlash);
            layout.Controls.Add(buttonPanel, 0, row++);

            timeSlotPanel.Controls.Add(layout);
            UpdateEveningEditors();
        }
        private void ShowTimeSlotPage() => contentPanel.Controls.Add(timeSlotPanel);

        // ---------- 集控管理页面 ----------
        private Panel mgmtPanel;
        private void CreateMgmtPage()
        {
            mgmtPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent, AutoSize = true };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            int row = 0;
            AddLabelAndControl(layout, "启用集控:", chkMgmtEnabled = new CheckBox { Text = "从远程服务器同步作业数据", Checked = config.MgmtEnabled, AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent }, row++);
            AddLabelAndControl(layout, "清单 URL:", txtMgmtManifestUrl = new TextBox { Text = config.MgmtManifestUrl, Width = 300, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }, row++);

            btnTestMgmtConnection = new Button { Text = "测试连接", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnTestMgmtConnection.Click += BtnTestMgmtConnection_Click;
            lblMgmtStatus = new Label { Text = "未测试", AutoSize = true, ForeColor = Color.LightGray, BackColor = Color.Transparent };
            FlowLayoutPanel testPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            testPanel.Controls.Add(btnTestMgmtConnection);
            testPanel.Controls.Add(lblMgmtStatus);
            layout.Controls.Add(testPanel, 1, row++);
            layout.SetRow(testPanel, row - 1);

            AddLabelAndControl(layout, "组织名称:", lblMgmtOrgName = new Label { Text = config.OrganizationName, AutoSize = true, ForeColor = Color.LightGreen, BackColor = Color.Transparent, Font = new Font("微软雅黑", 9, FontStyle.Bold) }, row++);

            mgmtPanel.Controls.Add(layout);
        }
        private void ShowMgmtPage() => contentPanel.Controls.Add(mgmtPanel);

        // ---------- 关于页面 ----------
        private Panel aboutPanel;
        private void CreateAboutPage()
        {
            aboutPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent, AutoSize = true };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 1,
                Padding = new Padding(0),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            int row = 0;

            layout.Controls.Add(new Label { Text = "作业展板 版本 2.0.0", Font = new Font("微软雅黑", 12, FontStyle.Bold), AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent, Margin = new Padding(0, 0, 0, 5) }, 0, row++);
            layout.Controls.Add(new Label { Text = "作者: MaxSui 隋修梁", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent, Font = new Font("微软雅黑", 10), Margin = new Padding(0, 0, 0, 5) }, 0, row++);
            layout.Controls.Add(new Label { Text = "© 2026 MaxSui 保留部分权利\n本软件遵循GNU General Public License 3.0开源协议", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent, Margin = new Padding(0, 0, 0, 15) }, 0, row++);

            AddSeparator(layout, ref row, true);

            // 高级选项（包含 WebView2 开关）
            AddSectionTitle(layout, "高级选项", ref row, true);
            chkUseWebView2 = new CheckBox
            {
                Text = "启用 WebView2 渲染（Markdown + LaTeX）",
                Checked = config.UseWebView2,
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 5, 0, 10)
            };
            layout.Controls.Add(chkUseWebView2, 0, row++);

            AddSeparator(layout, ref row, true);

            // 镜像站
            AddSectionTitle(layout, "镜像站", ref row, true);
            layout.Controls.Add(lblMirrorStatus = new Label
            {
                Text = "正在检测镜像站...",
                AutoSize = true,
                ForeColor = Color.LightGreen,
                BackColor = Color.Transparent,
                Font = new Font("微软雅黑", 9),
                Margin = new Padding(0, 5, 0, 5)
            }, 0, row++);
            layout.Controls.Add(cmbMirrorManual = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 200,
                Height = 25,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 0, 10)
            }, 0, row++);
            cmbMirrorManual.SelectedIndexChanged += async (s, e) =>
            {
                if (cmbMirrorManual.SelectedItem != null && cmbMirrorManual.SelectedIndex > 0)
                {
                    _mirrorManager.ClearCache();
                    await UpdateMirrorStatusAsync(lblMirrorStatus);
                }
            };

            AddSeparator(layout, ref row, true);

            // 更新
            AddSectionTitle(layout, "更新", ref row, true);
            layout.Controls.Add(progressDownload = new ProgressBar
            {
                Width = 200,
                Height = 20,
                Visible = false,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Margin = new Padding(0, 5, 0, 5)
            }, 0, row++);
            layout.Controls.Add(lblDownloadStatus = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Font = new Font("微软雅黑", 9),
                Visible = false
            }, 0, row++);
            FlowLayoutPanel updatePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent, Margin = new Padding(0, 10, 0, 10) };
            btnCheckUpdate = new Button { Text = "检查更新", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnCheckUpdate.Click += btnCheckUpdate_Click;
            btnSkipVersion = new Button { Text = "跳过", Width = 100, Height = 30, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Visible = false };
            btnSkipVersion.Click += (s, e) => { btnSkipVersion.Visible = false; btnCheckUpdate.Text = "检查更新"; };
            updatePanel.Controls.Add(btnCheckUpdate);
            updatePanel.Controls.Add(btnSkipVersion);
            layout.Controls.Add(updatePanel, 0, row++);
            layout.Controls.Add(lblUpdateContent = new Label { Text = "", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent, Font = new Font("微软雅黑", 9) }, 0, row++);

            AddSeparator(layout, ref row, true);

            // 重置
            AddSectionTitle(layout, "重置", ref row, true);
            btnResetAll = new Button
            {
                Text = "恢复所有设置为默认",
                Width = 150,
                Height = 30,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 10, 0, 0)
            };
            btnResetAll.Click += BtnResetAll_Click;
            layout.Controls.Add(btnResetAll, 0, row++);

            aboutPanel.Controls.Add(layout);
        }
        private void ShowAboutPage() => contentPanel.Controls.Add(aboutPanel);

        // ---------- 辅助布局方法 ----------
        private void AddLabelAndControl(TableLayoutPanel panel, string labelText, Control control, int row)
        {
            Label lbl = new Label
            {
                Text = labelText,
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = true,
                Margin = new Padding(0, 5, 10, 5)
            };
            panel.Controls.Add(lbl, 0, row);
            control.Margin = new Padding(0, 5, 0, 5);
            panel.Controls.Add(control, 1, row);
        }

        private void AddSectionTitle(TableLayoutPanel panel, string title, ref int row, bool singleColumn = false)
        {
            TableLayoutPanel titleLayout = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 1,
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 10, 0, 5)
            };
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            Label titleLabel = new Label
            {
                Text = title,
                Font = new Font("微软雅黑", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(200, 200, 200),
                BackColor = Color.Transparent,
                AutoSize = true,
                Margin = new Padding(0, 0, 10, 0)
            };

            LineControl line = new LineControl
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, titleLabel.Height / 2, 0, 0)
            };

            titleLayout.Controls.Add(titleLabel, 0, 0);
            titleLayout.Controls.Add(line, 1, 0);

            if (singleColumn)
            {
                panel.Controls.Add(titleLayout, 0, row);
                panel.SetColumnSpan(titleLayout, 1);
            }
            else
            {
                panel.Controls.Add(titleLayout, 0, row);
                panel.SetColumnSpan(titleLayout, 2);
            }
            row++;
        }

        private void AddSeparator(TableLayoutPanel panel, ref int row, bool singleColumn = false)
        {
            Panel line = new Panel
            {
                Height = 1,
                BackColor = Color.FromArgb(80, 80, 80),
                Margin = new Padding(0, 5, 0, 5)
            };
            if (singleColumn)
            {
                panel.Controls.Add(line, 0, row);
                panel.SetColumnSpan(line, 1);
            }
            else
            {
                panel.Controls.Add(line, 0, row);
                panel.SetColumnSpan(line, 2);
            }
            row++;
        }

        private ComboBox CreateComboBox(int width, string[] items)
        {
            var combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = width,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            combo.Items.AddRange(items);
            if (combo.Items.Count > 0)
                combo.SelectedIndex = 0;
            return combo;
        }

        private Control CreateColorPicker()
        {
            btnBarColor = new Button { Text = "选择颜色", Width = 80, Height = 25, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnBarColor.Click += BtnBarColor_Click;
            pnlBarColorPreview = new Panel { Width = 40, Height = 25, BackColor = Color.Yellow, BorderStyle = BorderStyle.FixedSingle };
            FlowLayoutPanel panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            panel.Controls.Add(btnBarColor);
            panel.Controls.Add(pnlBarColorPreview);
            return panel;
        }

        private Control CreateFontColorRadio()
        {
            FlowLayoutPanel panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            rbBlack = new RadioButton { Text = "黑色", Checked = true, AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            rbWhite = new RadioButton { Text = "白色", AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            panel.Controls.Add(rbBlack);
            panel.Controls.Add(rbWhite);
            return panel;
        }

        private Control CreateFontFamilySelector()
        {
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
            FlowLayoutPanel panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            panel.Controls.Add(cmbFontFamily);
            panel.Controls.Add(btnSelectCustomFont);
            return panel;
        }

        private Control CreateBgTypeRadio()
        {
            FlowLayoutPanel panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            rbTransparentBg = new RadioButton { Text = "透明背景", Checked = !config.UseBackgroundImage, AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            rbImageBg = new RadioButton { Text = "图片背景", Checked = config.UseBackgroundImage, AutoSize = true, ForeColor = Color.White, BackColor = Color.Transparent };
            panel.Controls.Add(rbTransparentBg);
            panel.Controls.Add(rbImageBg);
            Action updateVisibility = () =>
            {
                bool useImage = rbImageBg.Checked;
                txtBgImagePath.Visible = useImage;
                btnBrowseBgImage.Visible = useImage;
                cmbBgEffect.Enabled = !useImage;
            };
            rbTransparentBg.CheckedChanged += (s, e) => updateVisibility();
            rbImageBg.CheckedChanged += (s, e) => updateVisibility();
            return panel;
        }

        private Control CreateImagePathSelector()
        {
            txtBgImagePath = new TextBox
            {
                Width = 200,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Text = config.BackgroundImagePath
            };
            btnBrowseBgImage = new Button
            {
                Text = "浏览",
                Width = 60,
                Height = 23,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnBrowseBgImage.Click += (s, e) =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp";
                    ofd.Title = "选择背景图片";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        txtBgImagePath.Text = ofd.FileName;
                    }
                }
            };
            FlowLayoutPanel panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, BackColor = Color.Transparent };
            panel.Controls.Add(txtBgImagePath);
            panel.Controls.Add(btnBrowseBgImage);
            return panel;
        }

        // ---------- 事件处理 ----------
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

        private async void BtnTestMgmtConnection_Click(object sender, EventArgs e)
        {
            string url = txtMgmtManifestUrl.Text.Trim();
            if (string.IsNullOrEmpty(url))
            {
                lblMgmtStatus.Text = "❌ 请输入清单 URL";
                return;
            }

            btnTestMgmtConnection.Enabled = false;
            lblMgmtStatus.Text = "⏳ 测试中...";
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    string json = await client.GetStringAsync(url);
                    var manifest = JsonSerializer.Deserialize<Manifest>(json);
                    if (manifest != null)
                    {
                        lblMgmtStatus.Text = "✅ 连接成功";
                        lblMgmtOrgName.Text = manifest.OrganizationName;
                        config.OrganizationName = manifest.OrganizationName;
                    }
                    else
                    {
                        lblMgmtStatus.Text = "❌ 清单格式无效";
                    }
                }
            }
            catch (Exception ex)
            {
                lblMgmtStatus.Text = $"❌ 连接失败: {ex.Message}";
            }
            finally
            {
                btnTestMgmtConnection.Enabled = true;
            }
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

        private void BtnResetAll_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("确定要将所有设置恢复为默认值吗？此操作不可撤销。", "确认恢复", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                string configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Homework", "config.json");
                try
                {
                    if (File.Exists(configPath))
                        File.Delete(configPath);
                }
                catch { }

                config = AppConfig.Load();
                LoadSettings();
                mainForm.ApplySettings(config);
                MessageBox.Show("所有设置已恢复为默认值。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // ---------- 加载与保存 ----------
        private void LoadSettings()
        {
            cmbMode.SelectedItem = config.LastMode;
            if (cmbMode.SelectedIndex == -1 && cmbMode.Items.Count > 0) cmbMode.SelectedIndex = 0;

            if (config.FontSizeLevel >= 0 && config.FontSizeLevel < cmbFontSize.Items.Count)
                cmbFontSize.SelectedIndex = config.FontSizeLevel;
            else if (cmbFontSize.Items.Count > 0)
                cmbFontSize.SelectedIndex = 0;

            trackCardOpacity.Value = config.CardOpacity;
            numCardOpacity.Value = config.CardOpacity;
            trackBgOpacity.Value = config.BackgroundOpacity;
            numBgOpacity.Value = config.BackgroundOpacity;
            if (config.FontColorWhite) rbWhite.Checked = true; else rbBlack.Checked = true;
            pnlBarColorPreview.BackColor = ParseColor(config.BarColor, Color.Yellow);
            numScrollSpeed.Value = config.ScrollSpeed;
            cmbDefaultExportFormat.SelectedItem = config.ExportFormat;
            cmbFontFamily.SelectedItem = config.IsCustomFont ? "自定义" : config.FontFamily;
            chkShowMouseGlow.Checked = config.ShowMouseGlow;
            chkShowDueTime.Checked = config.ShowDueTime;
            chkMgmtEnabled.Checked = config.MgmtEnabled;
            txtMgmtManifestUrl.Text = config.MgmtManifestUrl;
            lblMgmtOrgName.Text = config.OrganizationName;
            chkUseWebView2.Checked = config.UseWebView2;

            rbTransparentBg.Checked = !config.UseBackgroundImage;
            rbImageBg.Checked = config.UseBackgroundImage;
            txtBgImagePath.Text = config.BackgroundImagePath;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            bool oldUseWebView2 = config.UseWebView2;

            config.LastMode = cmbMode.SelectedItem?.ToString() ?? "大理";
            config.FontSizeLevel = cmbFontSize.SelectedIndex;
            config.CardOpacity = trackCardOpacity.Value;
            config.BackgroundOpacity = trackBgOpacity.Value;
            config.FontColorWhite = rbWhite.Checked;
            config.ScrollSpeed = (int)numScrollSpeed.Value;
            config.BackgroundEffect = cmbBgEffect.SelectedItem?.ToString() ?? "Mica";
            config.ExportFormat = cmbDefaultExportFormat.SelectedItem?.ToString() ?? "txt";
            config.ShowMouseGlow = chkShowMouseGlow.Checked;
            config.ShowDueTime = chkShowDueTime.Checked;

            string fontSelection = cmbFontFamily.SelectedItem?.ToString();
            if (fontSelection == "自定义")
                config.IsCustomFont = true;
            else
            {
                config.IsCustomFont = false;
                config.FontFamily = fontSelection;
            }

            config.MgmtEnabled = chkMgmtEnabled.Checked;
            config.MgmtManifestUrl = txtMgmtManifestUrl.Text.Trim();

            config.EveningClassCount = (int)numEveningCount.Value;
            config.EveningClassTimes.Clear();
            for (int i = 0; i < startPickers.Count; i++)
                config.EveningClassTimes.Add(new EveningClassTime { Start = startPickers[i].Value.ToString("HH:mm"), End = endPickers[i].Value.ToString("HH:mm") });

            string bgPath = txtBgImagePath.Text.Trim();
            if (!string.IsNullOrEmpty(bgPath) && !Path.IsPathRooted(bgPath))
                bgPath = Path.Combine(Application.StartupPath, bgPath);
            config.UseBackgroundImage = rbImageBg.Checked;
            config.BackgroundImagePath = bgPath;

            // 新增 WebView2 选项
            config.UseWebView2 = chkUseWebView2.Checked;

            config.Save();

            mainForm.ApplySettings(config);
            this.DialogResult = DialogResult.OK;

            // 如果 WebView2 使用状态发生变化，提示重启
            if (config.UseWebView2 != oldUseWebView2)
            {
                MessageBox.Show("WebView2 渲染设置已更改，请重启应用以生效。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            this.Close();
        }
    }
}