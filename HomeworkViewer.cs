#nullable disable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Timer = System.Windows.Forms.Timer;

namespace HomeworkViewer
{
    public class HomeworkViewer : Form
    {
        // ====================== 常量 ======================
        private readonly Size VIRTUAL_SIZE = new Size(1200, 675);
        private const int BTN_SQUARE_SIZE = 46;
        private readonly Point FULLSCREEN_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10, 13);
        private readonly Point HISTORY_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10, 13);
        private readonly Point EDIT_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10, 13);
        private readonly Point ROTATE_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10, 13);
        private Point SETTINGS_BTN_POS;
        private readonly Rectangle GRID_AREA = new Rectangle(0, 47, 1200, 631);
        private const int GRID_PADDING = 20, ROUND_RADIUS = 5;
        private const int BAR_HEIGHT = 46, BAR_Y = 0, BAR_PADDING = 20;

        // 顶部下拉栏
        private Panel _drawerPanel;
        private bool _drawerOpen = false;
        private const int MAX_DRAWER_HEIGHT_RATIO = 3;

        // 动画
        private Timer _animationTimer;
        private List<Animation> _animations = new List<Animation>();
        private Dictionary<int, Animation> _cardAnimations = new Dictionary<int, Animation>();
        private Animation _fadeAnimation;
        private string _oldRotationText = "";

        // 轮播滑动动画（已取消）
        private bool _isSliding = false;
        private float _slideProgress = 0;
        private int _slidePhase = 0;
        private int _oldSlideIndex = -1;
        private int _newSlideIndex = -1;
        private string _oldSlideContent = "";
        private string _newSlideContent = "";

        // Mica/亚克力
        private bool _micaEnabled = false;
        private int _backgroundAlpha = 32;
        private Image _backgroundImage = null;

        // 状态
        private bool fullscreen = false;
        private bool historyMode = false;
        private bool editMode = false;
        private bool rotationMode = false;
        private int rotationIndex = 0;
        private Timer rotationTimer;
        private const int rotationInterval = 10000;

        // 内联编辑（已弃用）
        private Control inlineEditControl;
        private int editingSubjectIndex = -1;
        private enum EditFieldType { None, Subject, DueTime }
        private EditFieldType currentEditType = EditFieldType.None;

        // 数据
        private HomeworkData homeworkData = new HomeworkData();
        private DateTime currentDate = DateTime.Now;
        private DateTime? historyDate = null;

        // 按钮矩形
        private Rectangle rotateBtnRect, editBtnRect, historyBtnRect, fullscreenBtnRect, backBtnRect, leftArrowRect, rightArrowRect, settingsBtnRect, minimizeBtnRect, closeBtnRect, exportBtnRect;
        private Rectangle expandBtnRect, collapseBtnRect;

        // 网格
        private List<Rectangle> gridRects = new List<Rectangle>();

        // 字体
        private Font font12, font20, font24, font30, font22, font36, hintFont, buttonFont, fontSmall;

        // 颜色
        private Color TEXT_COLOR = Color.Black;
        private readonly Brush RED_SEMI = new SolidBrush(Color.FromArgb(255, 255, 0, 0));
        private readonly Brush ORANGE_SEMI = new SolidBrush(Color.FromArgb(255, 255, 165, 0));
        private readonly Brush GREEN_SEMI = new SolidBrush(Color.FromArgb(255, 0, 255, 0));
        private readonly Brush BLUE_SEMI = new SolidBrush(Color.FromArgb(255, 0, 0, 255));
        private readonly Brush PURPLE_SEMI = new SolidBrush(Color.FromArgb(255, 128, 0, 128));
        private readonly Brush DARKORANGE_SEMI = new SolidBrush(Color.FromArgb(255, 255, 140, 0));

        // 缩放
        private float scaleFactor = 1.0f;
        private Point offset = Point.Empty;
        private bool _resizing = false;
        private bool _inSizeMove = false;

        // 图片资源
        private Image buttonImage, historyBtnImage, backBtnImage, leftArrowImage, rightArrowImage;
        private Dictionary<string, Image> buttonIcons = new Dictionary<string, Image>();

        // 配置
        private AppConfig appConfig;
        private float[] fontScales = { 0.8f, 1.0f, 1.2f };
        private string[] currentSubjects;
        private bool _useWebView2 = false;

        // 窗口状态保存
        private Size _savedClientSize;
        private FormWindowState _savedWindowState;

        // 按钮悬停/按下
        private Dictionary<string, bool> _buttonHover = new Dictionary<string, bool>();
        private Dictionary<string, bool> _buttonPressed = new Dictionary<string, bool>();
        private Dictionary<string, float> _buttonHoverProgress = new Dictionary<string, float>();
        private string _currentHoverKey = null;

        // 按钮按下标志
        private bool _settingsPressed = false, _rotatePressed = false, _editPressed = false, _historyPressed = false, _fullscreenPressed = false, _backPressed = false, _minimizePressed = false, _closePressed = false, _expandPressed = false, _collapsePressed = false, _exportPressed = false;

        // 收纳
        private bool _buttonsExpanded = false;

        // 模式切换下拉框
        private ComboBox modeComboBox;
        private List<ComboBox> timeComboBoxes = new List<ComboBox>();

        // 时间提醒
        private Timer timeCheckTimer;
        private List<string> _activeEvenings = new List<string>();
        private List<string> _flashingEvenings = new List<string>();
        private List<string> _grayEvenings = new List<string>();
        private bool _previousFlashingState = false;

        // 闪烁
        private Timer flashTimer;
        private int flashStep = 0;
        private bool _debugFlashing = false;
        private DateTime _debugFlashStartTime;
        private DateTime flashStartTime;
        private const int FLASH_DURATION = 300;
        private const int FLASH_INTERVAL = 100;
        private float _laserOffset = 0f;

        // 滚动（用于 GDI+ 模式）
        private Timer scrollTimer;
        private Dictionary<int, float> scrollOffsets = new Dictionary<int, float>();
        private Dictionary<int, bool> scrollPaused = new Dictionary<int, bool>();
        private Dictionary<int, DateTime> pauseStartTime = new Dictionary<int, DateTime>();
        private const int SCROLL_PAUSE_SECONDS = 3;

        // 背景效果
        private string _currentBackgroundEffect = "Mica";
        private bool _isWin10OrAbove = false;
        private bool _isWin10 = false;
        private bool _sizing = false;

        // 未使用的 Markdown 滚动字段（兼容旧代码）
        private float _markdownScrollOffset = 0f;
        private bool _markdownScrollPaused = false;
        private DateTime _markdownPauseStart = DateTime.Now;
        private float _markdownTotalHeight = 0;

        // 鼠标跟随光晕坐标
        private Point _mouseVirtualPos = new Point(-1000, -1000);
        private bool _hideMouseGlow = false;
        private bool _suspendMouseGlow = false;
        private bool _isFlyingIn = false;

        // 光晕透明度控制
        private float _mouseGlowAlpha = 1.0f;
        private Animation _mouseGlowFadeAnimation;
        private Animation _mouseGlowFadeOutAnimation;
        private bool _mouseInside = true;
        private int _flyInAnimationsRemaining = 0;

        // 涟漪效果
        private Point _rippleCenter = new Point(-1, -1);
        private float _rippleRadius = 0;
        private float _rippleAlpha = 0;
        private Animation _rippleAnimation;
        private bool _inRipple = false;

        // ========== WebView2 相关 ==========
        private List<WebView2> cardWebViews = new List<WebView2>();
        private WebView2 rotationWebView;
        private bool webView2Initialized = false;

        // ========== 编辑模式下的纯文本编辑框（两种模式共用） ==========
        private List<TextBox> plainTextEditors = new List<TextBox>();

        // ========== Markdown 渲染 HTML 模板（仅 WebView2 模式使用） ==========
        private const string MarkdownHtmlTemplate = @"<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <style>
        body {
            font-family: 'Microsoft YaHei', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
            margin: 12px;
            padding: 0;
            background-color: transparent;
            color: {textColor} !important;
            font-size: {fontSize}px;
            line-height: 1.5;
            overflow-y: auto;
            overflow-x: hidden;
        }
        body * { color: {textColor} !important; }
        pre { background-color: #2d2d2d; padding: 8px; border-radius: 6px; overflow-x: auto; }
        code { font-family: 'Consolas', 'Courier New', monospace; background-color: rgba(0,0,0,0.05); padding: 2px 4px; border-radius: 4px; }
        blockquote { border-left: 4px solid #ccc; margin: 0; padding-left: 16px; color: #666; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 6px; text-align: left; }
        th { background-color: #f2f2f2; }
        a { color: #0366d6; text-decoration: none; }
        a:hover { text-decoration: underline; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #888; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #555; }
    </style>
    <link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/katex.min.css"">
    <script src=""https://cdn.jsdelivr.net/npm/markdown-it@13.0.1/dist/markdown-it.min.js""></script>
    <script src=""https://cdn.jsdelivr.net/npm/markdown-it-texmath@1.0.0/texmath.js""></script>
    <script src=""https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/katex.min.js""></script>
    <script>
        var md = window.markdownit({
            html: true,
            breaks: true,
            linkify: true
        }).use(texmath, {
            engine: katex,
            delimiters: ['dollars', 'brackets'],
            katexOptions: { macros: {} }
        });
        function render() {
            var raw = document.getElementById('raw').innerText;
            var result = md.render(raw);
            document.getElementById('content').innerHTML = result;
        }
    </script>
</head>
<body>
    <pre id=""raw"" style=""display:none;"">{markdownText}</pre>
    <div id=""content""></div>
    <script>render();</script>
</body>
</html>";

        // ------------------------------ 构造函数 ------------------------------
        public HomeworkViewer()
        {
            SETTINGS_BTN_POS = new Point(ROTATE_BTN_POS.X - BTN_SQUARE_SIZE - 10, 13);

            Text = "作业展板";
            ClientSize = VIRTUAL_SIZE;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;
            DoubleBuffered = true;
            KeyPreview = true;
            BackColor = Color.Black;
            ResizeRedraw = false;

            appConfig = AppConfig.Load();
            _useWebView2 = appConfig.UseWebView2;
            _backgroundAlpha = (int)(appConfig.BackgroundOpacity / 100f * 255);
            ApplyFontSettings();
            UpdateSubjectsByMode();
            CalculateGridRects();
            LoadHomeworkData(currentDate);
            LoadImages();
            LoadBackgroundImage();

            rotationTimer = new Timer { Interval = rotationInterval };
            rotationTimer.Tick += (s, e) => { if (rotationMode) RotateNext(); };

            timeCheckTimer = new Timer { Interval = 1000 };
            timeCheckTimer.Tick += (s, e) => CheckEveningClassStates();
            timeCheckTimer.Start();

            flashTimer = new Timer { Interval = FLASH_INTERVAL };
            flashTimer.Tick += FlashTimer_Tick;

            scrollTimer = new Timer { Interval = 50 };
            scrollTimer.Tick += ScrollTimer_Tick;
            scrollTimer.Start();

            _animationTimer = new Timer { Interval = 16 };
            _animationTimer.Tick += (s, e) => {
                bool needRedraw = false;
                foreach (var anim in _animations.ToList())
                {
                    anim.Update();
                    needRedraw = true;
                }
                if (needRedraw) Invalidate();
                _animations.RemoveAll(a => !a.IsRunning);
            };
            _animationTimer.Start();

            this.MouseClick += OnMouseClick;
            this.MouseDoubleClick += OnMouseDoubleClick;
            this.MouseDown += OnMouseDown;
            this.MouseUp += OnMouseUp;
            this.MouseLeave += OnMouseLeave;
            this.MouseMove += OnMouseMove;
            this.Resize += OnResize;
            this.Load += OnLoad;
            this.Activated += OnActivated;
            this.FormClosing += OnFormClosing;
            this.KeyDown += OnKeyDown;

            InitializeModeComboBox();

            CheckWindowsVersion();
            ApplyBackgroundEffect(appConfig.BackgroundEffect);
            CheckEveningClassStates();

            ManagementHelper.CheckForUpdatesAsync(appConfig, (updated) => { if (updated) Invalidate(); });

            if (_useWebView2)
                InitializeWebView2Async();
            else
                webView2Initialized = false;
        }

        // ------------------------------ WebView2 初始化 ------------------------------
        private async void InitializeWebView2Async()
        {
            if (!_useWebView2) return;
            try
            {
                var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync();
                webView2Initialized = true;
                if (!rotationMode && !EditMode && !IsDisposed)
                {
                    this.Invoke((MethodInvoker)(() => RefreshAllWebView2Content()));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"WebView2 初始化失败: {ex.Message}\n\nMarkdown 渲染将不可用，请安装 WebView2 运行时。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ------------------------------ EditMode 属性 ------------------------------
        private bool EditMode
        {
            get => editMode;
            set
            {
                if (editMode != value)
                {
                    editMode = value;
                    if (editMode)
                    {
                        // 进入编辑模式：销毁只读控件（WebView2 或 GDI+ 无需销毁，但需隐藏文本绘制），创建纯文本编辑框
                        if (_useWebView2)
                            DestroyAllWebView2();
                        // 注意：GDI+ 模式不需要销毁控件，只需在 DrawGridBackground 中跳过绘制文本（通过 editMode 标志控制）
                        CreateTimeComboBoxes();
                        CreatePlainTextEditors();
                        scrollOffsets.Clear();
                        scrollPaused.Clear();
                        pauseStartTime.Clear();
                    }
                    else
                    {
                        // 退出编辑模式：销毁纯文本编辑框和时间下拉框，根据模式重建只读视图
                        DestroyPlainTextEditors();
                        DestroyTimeComboBoxes();
                        if (_useWebView2)
                        {
                            RecreateWebView2ForGrid();
                            RefreshAllWebView2Content();
                        }
                        // GDI+ 模式无需重建，只需重绘即可
                        scrollOffsets.Clear();
                        scrollPaused.Clear();
                        pauseStartTime.Clear();
                        _markdownScrollOffset = 0f;
                        _markdownScrollPaused = false;
                        _markdownPauseStart = DateTime.Now;
                        _markdownTotalHeight = 0;
                    }
                    Invalidate();
                }
            }
        }

        // ------------------------------ 纯文本编辑框管理（编辑模式共用） ------------------------------
        private void CreatePlainTextEditors()
        {
            DestroyPlainTextEditors();
            for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
            {
                string subject = currentSubjects[i];
                string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
                Rectangle textArea = GetContentArea(i);
                Point screenLoc = MapToScreen(textArea.Location);
                Size screenSize = new Size((int)(textArea.Width * scaleFactor), (int)(textArea.Height * scaleFactor));

                var textBox = new TextBox
                {
                    Multiline = true,
                    WordWrap = true,
                    ScrollBars = ScrollBars.Vertical,
                    Location = screenLoc,
                    Size = screenSize,
                    Text = content,
                    Font = font30,
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(40, 40, 40),
                    ForeColor = Color.White,
                    Tag = i
                };
                textBox.TextChanged += PlainTextBox_TextChanged;
                textBox.LostFocus += PlainTextBox_LostFocus;
                Controls.Add(textBox);
                plainTextEditors.Add(textBox);
            }
        }

        private void DestroyPlainTextEditors()
        {
            foreach (var tb in plainTextEditors)
            {
                if (tb != null && !tb.IsDisposed)
                {
                    tb.TextChanged -= PlainTextBox_TextChanged;
                    tb.LostFocus -= PlainTextBox_LostFocus;
                    Controls.Remove(tb);
                    tb.Dispose();
                }
            }
            plainTextEditors.Clear();
        }

        private void PlainTextBox_TextChanged(object sender, EventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;
            int index = (int)tb.Tag;
            if (index >= 0 && index < currentSubjects.Length)
            {
                string subject = currentSubjects[index];
                homeworkData.Subjects[subject] = tb.Text;
            }
        }

        private void PlainTextBox_LostFocus(object sender, EventArgs e)
        {
            SaveHomeworkData();
        }

        // ------------------------------ WebView2 管理（仅 _useWebView2 时使用） ------------------------------
        private void CreateWebView2ForCard(int index, Rectangle virtualRect)
        {
            if (!_useWebView2 || !webView2Initialized) return;
            if (index >= cardWebViews.Count)
            {
                var webView = new WebView2();
                webView.Visible = false;
                webView.DefaultBackgroundColor = Color.Transparent;
                this.Controls.Add(webView);
                cardWebViews.Add(webView);
            }
            var wv = cardWebViews[index];
            wv.Visible = !editMode && !rotationMode;
            UpdateWebView2Position(wv, virtualRect);
            string subject = currentSubjects[index];
            string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
            _ = UpdateWebView2Content(wv, content);
        }

        private async Task UpdateWebView2Content(WebView2 webView, string markdownText)
        {
            if (webView == null || !webView2Initialized) return;
            if (string.IsNullOrWhiteSpace(markdownText))
                markdownText = "*（无作业内容）*";

            string textColorHex = appConfig.FontColorWhite ? "#FFFFFF" : "#000000";
            int fontSize = (int)(font30.SizeInPoints * 1.2f);
            string html = MarkdownHtmlTemplate
                .Replace("{textColor}", textColorHex)
                .Replace("{fontSize}", fontSize.ToString())
                .Replace("{markdownText}", markdownText.Replace("\\", "\\\\").Replace("`", "\\`").Replace("$", "\\$"));
            try
            {
                await webView.EnsureCoreWebView2Async();
                webView.CoreWebView2.NavigateToString(html);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebView2 导航失败: {ex.Message}");
            }
        }

        private void UpdateWebView2Position(WebView2 webView, Rectangle virtualRect)
        {
            if (webView == null) return;
            Point screenLoc = MapToScreen(virtualRect.Location);
            Size screenSize = new Size((int)(virtualRect.Width * scaleFactor), (int)(virtualRect.Height * scaleFactor));
            webView.Location = screenLoc;
            webView.Size = screenSize;
        }

        private void UpdateWebView2Positions()
        {
            if (rotationMode && rotationWebView != null)
            {
                Rectangle webViewRect = new Rectangle(150, 230, 900, 345);
                UpdateWebView2Position(rotationWebView, webViewRect);
            }
            else if (!EditMode && !rotationMode && _useWebView2)
            {
                for (int i = 0; i < gridRects.Count && i < cardWebViews.Count; i++)
                {
                    Rectangle textArea = GetContentArea(i);
                    UpdateWebView2Position(cardWebViews[i], textArea);
                }
            }
        }

        private void DestroyAllWebView2()
        {
            foreach (var wv in cardWebViews)
            {
                if (wv != null && !wv.IsDisposed)
                    wv.Dispose();
            }
            cardWebViews.Clear();
            if (rotationWebView != null && !rotationWebView.IsDisposed)
            {
                rotationWebView.Dispose();
                rotationWebView = null;
            }
        }

        private void RecreateWebView2ForGrid()
        {
            if (!_useWebView2) return;
            DestroyAllWebView2();
            for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
            {
                Rectangle textArea = GetContentArea(i);
                CreateWebView2ForCard(i, textArea);
            }
        }

        private async void RefreshAllWebView2Content()
        {
            if (!_useWebView2 || !webView2Initialized) return;
            for (int i = 0; i < currentSubjects.Length && i < cardWebViews.Count; i++)
            {
                string subject = currentSubjects[i];
                string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
                await UpdateWebView2Content(cardWebViews[i], content);
            }
            if (rotationWebView != null && rotationMode && rotationIndex < currentSubjects.Length)
            {
                string subject = currentSubjects[rotationIndex];
                string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
                await UpdateWebView2Content(rotationWebView, content);
            }
        }

        private void CreateRotationWebView()
        {
            if (!_useWebView2 || !webView2Initialized) return;
            if (rotationWebView == null)
            {
                rotationWebView = new WebView2();
                rotationWebView.DefaultBackgroundColor = Color.Transparent;
                this.Controls.Add(rotationWebView);
            }
            rotationWebView.Visible = true;
            Rectangle webViewRect = new Rectangle(150, 230, 900, 345);
            UpdateWebView2Position(rotationWebView, webViewRect);
            string subject = currentSubjects[rotationIndex];
            string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
            _ = UpdateWebView2Content(rotationWebView, content);
        }

        // ------------------------------ 辅助方法 ------------------------------
        private class BufferedPanel : Panel
        {
            public BufferedPanel()
            {
                this.DoubleBuffered = true;
                this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
            }
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.I)
            {
                using (var settingsForm = new SettingsForm(this))
                {
                    settingsForm.ShowDialog();
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F11)
            {
                ToggleFullscreen();
                e.Handled = true;
            }
        }

        private void CheckWindowsVersion()
        {
            var os = Environment.OSVersion;
            _isWin10OrAbove = os.Version.Major >= 10;
            _isWin10 = os.Version.Major == 10 && os.Version.Build < 22000;
        }

        private void LoadBackgroundImage()
        {
            if (appConfig.UseBackgroundImage && !string.IsNullOrEmpty(appConfig.BackgroundImagePath))
            {
                string fullPath = appConfig.BackgroundImagePath;
                if (!Path.IsPathRooted(fullPath))
                    fullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fullPath);
                if (File.Exists(fullPath))
                {
                    try
                    {
                        _backgroundImage?.Dispose();
                        _backgroundImage = Image.FromFile(fullPath);
                    }
                    catch { _backgroundImage = null; }
                }
                else _backgroundImage = null;
            }
            else _backgroundImage = null;
        }

        public void ApplyBackgroundEffect(string effect)
        {
            _currentBackgroundEffect = effect;
            if (appConfig.UseBackgroundImage)
            {
                _micaEnabled = false;
                Invalidate();
                return;
            }
            try
            {
                switch (effect)
                {
                    case "Mica": EnableMica(); break;
                    case "Acrylic": EnableAcrylic(); break;
                    case "Aero": EnableAero(); break;
                    default: EnableMica(); break;
                }
            }
            catch { }
        }

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_MICA = 1029;
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private enum AccentState
        {
            ACCENT_DISABLED = 0,
            ACCENT_ENABLE_GRADIENT = 1,
            ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
            ACCENT_ENABLE_BLURBEHIND = 3,
            ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct AccentPolicy
        {
            public AccentState AccentState;
            public uint AccentFlags;
            public uint GradientColor;
            public uint AnimationId;
        }

        private enum WindowCompositionAttribute
        {
            WCA_ACCENT_POLICY = 19
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WindowCompositionAttributeData
        {
            public WindowCompositionAttribute Attribute;
            public IntPtr Data;
            public int SizeOfData;
        }

        [DllImport("user32.dll")]
        private static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

        private void EnableMica()
        {
            try
            {
                if (Environment.OSVersion.Version.Build >= 22000)
                {
                    int micaValue = 1;
                    int result = DwmSetWindowAttribute(this.Handle, DWMWA_MICA, ref micaValue, sizeof(int));
                    if (result == 0)
                    {
                        _micaEnabled = true;
                        int darkMode = 0;
                        DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, sizeof(int));
                    }
                    else EnableAcrylicFallback();
                }
                else EnableAcrylicFallback();
            }
            catch { }
        }

        private void EnableAcrylic()
        {
            try
            {
                var accent = new AccentPolicy
                {
                    AccentState = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND,
                    GradientColor = 0x99000000
                };
                int accentStructSize = Marshal.SizeOf(accent);
                IntPtr accentPtr = Marshal.AllocHGlobal(accentStructSize);
                Marshal.StructureToPtr(accent, accentPtr, false);
                var data = new WindowCompositionAttributeData
                {
                    Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY,
                    SizeOfData = accentStructSize,
                    Data = accentPtr
                };
                SetWindowCompositionAttribute(this.Handle, ref data);
                Marshal.FreeHGlobal(accentPtr);
                _micaEnabled = true;
            }
            catch { }
        }

        private void EnableAero()
        {
            try
            {
                var accent = new AccentPolicy
                {
                    AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND,
                    GradientColor = 0x99000000
                };
                int accentStructSize = Marshal.SizeOf(accent);
                IntPtr accentPtr = Marshal.AllocHGlobal(accentStructSize);
                Marshal.StructureToPtr(accent, accentPtr, false);
                var data = new WindowCompositionAttributeData
                {
                    Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY,
                    SizeOfData = accentStructSize,
                    Data = accentPtr
                };
                SetWindowCompositionAttribute(this.Handle, ref data);
                Marshal.FreeHGlobal(accentPtr);
                _micaEnabled = true;
            }
            catch { }
        }

        private void EnableAcrylicFallback() => EnableAcrylic();

        protected override void WndProc(ref Message m)
        {
            const int WM_ENTERSIZEMOVE = 0x0231;
            const int WM_EXITSIZEMOVE = 0x0232;
            const int WM_SYSCOMMAND = 0x0112;
            const int SC_MAXIMIZE = 0xF030;
            const int WM_SIZE = 0x0005;
            const int SIZE_RESTORED = 0;

            if (_isWin10OrAbove)
            {
                if (m.Msg == WM_ENTERSIZEMOVE)
                {
                    _sizing = true;
                    _inSizeMove = true;
                    DestroyAllDynamicControls();
                    Invalidate();
                }
                else if (m.Msg == WM_EXITSIZEMOVE)
                {
                    _sizing = false;
                    _inSizeMove = false;
                    UpdateScale();
                    CalculateGridRects();
                    RecreateDynamicControls();
                    Invalidate(true);
                }
            }

            if (m.Msg == WM_SYSCOMMAND && (int)m.WParam == SC_MAXIMIZE)
            {
                ToggleFullscreen();
                return;
            }

            if (m.Msg == WM_SIZE && m.WParam.ToInt32() == SIZE_RESTORED)
            {
                if (_savedClientSize != Size.Empty)
                    this.ClientSize = _savedClientSize;
                UpdateScale();
                if (!_sizing) Invalidate();
                if (!_micaEnabled) ApplyBackgroundEffect(_currentBackgroundEffect);
            }

            base.WndProc(ref m);
        }

        private void OnResize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                _savedClientSize = ClientSize;
                _savedWindowState = WindowState;
                return;
            }

            if (_resizing || fullscreen || WindowState == FormWindowState.Maximized) return;

            _resizing = true;
            float targetRatio = (float)VIRTUAL_SIZE.Width / VIRTUAL_SIZE.Height;
            int newWidth = ClientSize.Width;
            int newHeight = (int)(newWidth / targetRatio);
            ClientSize = new Size(newWidth, newHeight);
            _resizing = false;

            UpdateScale();
            if (EditMode) CreateTimeComboBoxes();
            if (!_isWin10OrAbove || !_sizing) Invalidate();

            if (_drawerPanel != null)
            {
                _drawerPanel.Width = this.ClientSize.Width;
                if (_drawerOpen)
                {
                    int targetHeight = this.ClientSize.Height / MAX_DRAWER_HEIGHT_RATIO;
                    SetDrawerHeight(targetHeight);
                }
            }

            // 更新动态控件位置
            if (EditMode)
            {
                for (int i = 0; i < plainTextEditors.Count && i < gridRects.Count; i++)
                {
                    Rectangle textArea = GetContentArea(i);
                    Point screenLoc = MapToScreen(textArea.Location);
                    Size screenSize = new Size((int)(textArea.Width * scaleFactor), (int)(textArea.Height * scaleFactor));
                    plainTextEditors[i].Location = screenLoc;
                    plainTextEditors[i].Size = screenSize;
                }
            }
            else if (_useWebView2)
            {
                UpdateWebView2Positions();
            }
            // GDI+ 模式无需更新额外控件，文本绘制依赖 gridRects 重绘即可
        }

        public void ApplySettings(AppConfig newConfig)
        {
            if (newConfig.ColumnWidth > 2000 || newConfig.ColumnWidth < 0) newConfig.ColumnWidth = 0;
            if (newConfig.RowHeight > 1000 || newConfig.RowHeight < 0) newConfig.RowHeight = 0;

            bool oldUseWebView2 = _useWebView2;
            appConfig = newConfig;
            _useWebView2 = appConfig.UseWebView2;
            ApplyFontSettings();
            UpdateSubjectsByMode();
            _backgroundAlpha = (int)(appConfig.BackgroundOpacity / 100f * 255);
            CheckEveningClassStates();
            if (EditMode) CreateTimeComboBoxes();
            LoadBackgroundImage();
            ApplyBackgroundEffect(appConfig.BackgroundEffect);
            CalculateGridRects();

            if (oldUseWebView2 != _useWebView2)
            {
                MessageBox.Show("WebView2 渲染设置已更改，请重启应用以生效。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // 不动态切换，保持当前模式
            }
            else
            {
                if (_useWebView2)
                {
                    if (!EditMode && !rotationMode)
                    {
                        RecreateWebView2ForGrid();
                        RefreshAllWebView2Content();
                    }
                }
                // GDI+ 模式无需重建，只需重绘
            }
            Invalidate();
        }

        private void ApplyFontSettings()
        {
            font12?.Dispose(); font20?.Dispose(); font24?.Dispose(); font30?.Dispose();
            font22?.Dispose(); font36?.Dispose(); hintFont?.Dispose(); buttonFont?.Dispose(); fontSmall?.Dispose();

            float scale = fontScales[appConfig.FontSizeLevel];
            string fontName = appConfig.FontFamily;

            if (appConfig.IsCustomFont)
            {
                var customFont = FontManager.GetCustomFont(20 * scale);
                font30 = new Font(customFont.FontFamily, 20 * scale);
                font22 = new Font(customFont.FontFamily, 21 * scale, FontStyle.Bold);
                font12 = new Font(customFont.FontFamily, 10 * scale);
                font20 = new Font(customFont.FontFamily, 15 * scale);
                font24 = new Font(customFont.FontFamily, 14 * scale);
                font36 = new Font(customFont.FontFamily, 35 * scale);
                hintFont = new Font(customFont.FontFamily, 15 * scale);
                buttonFont = new Font(customFont.FontFamily, 10 * scale);
                fontSmall = new Font(customFont.FontFamily, 8 * scale);
                return;
            }

            font12 = new Font(fontName, 10 * scale);
            font20 = new Font(fontName, 15 * scale);
            font24 = new Font(fontName, 14 * scale);
            font30 = new Font(fontName, 20 * scale);
            font22 = new Font(fontName, 21 * scale, FontStyle.Bold);
            font36 = new Font(fontName, 35 * scale);
            hintFont = new Font(fontName, 15 * scale);
            buttonFont = new Font(fontName, 10 * scale);
            fontSmall = new Font(fontName, 8 * scale);

            TEXT_COLOR = appConfig.FontColorWhite ? Color.White : Color.Black;
        }

        private void UpdateSubjectsByMode()
        {
            var subjectModes = new Dictionary<string, string[]>
            {
                {"大理", new[]{"语文","数学","英语","物理","化学","生物"}},
                {"中理", new[]{"语文","数学","英语","物理","化学","地理"}},
                {"小理", new[]{"语文","数学","英语","物理","化学","政治"}},
                {"大文", new[]{"语文","数学","英语","政治","历史","地理"}},
                {"全科", new[]{"语文","数学","英语","物理","化学","生物","政治","历史","地理"}}
            };
            currentSubjects = subjectModes.ContainsKey(appConfig.LastMode) ? subjectModes[appConfig.LastMode] : subjectModes["大理"];
            CalculateGridRects();
        }

        private void CalculateGridRects()
        {
            gridRects.Clear();
            int subjectCount = currentSubjects.Length;
            int cols = 3;
            int rows = (subjectCount + cols - 1) / cols;

            float areaWidth = GRID_AREA.Width;
            float areaHeight = GRID_AREA.Height;

            float rectWidth = appConfig.ColumnWidth > 0 ? appConfig.ColumnWidth : (areaWidth - (cols + 1) * GRID_PADDING) / cols;
            float rectHeight = appConfig.RowHeight > 0 ? appConfig.RowHeight : (areaHeight - (rows + 1) * GRID_PADDING) / rows;

            for (int row = 0; row < rows; row++)
                for (int col = 0; col < cols; col++)
                {
                    int index = row * cols + col;
                    if (index >= subjectCount) break;
                    int x = GRID_AREA.Left + GRID_PADDING + col * (int)(rectWidth + GRID_PADDING);
                    int y = GRID_AREA.Top + GRID_PADDING + row * (int)(rectHeight + GRID_PADDING);
                    gridRects.Add(new Rectangle(x, y, (int)rectWidth, (int)rectHeight));
                }
            _cardAnimations.Clear();
        }

        private void LoadHomeworkData(DateTime date)
        {
            if (appConfig.MgmtEnabled && appConfig.MgmtForceRemote)
            {
                var remoteData = ManagementHelper.LoadHomeworkDataFromRemote(date);
                if (remoteData != null)
                {
                    homeworkData = remoteData;
                    if (EditMode) CreateTimeComboBoxes();
                    if (_useWebView2) RefreshAllWebView2Content();
                    else Invalidate(); // GDI+ 模式重绘
                    return;
                }
            }

            homeworkData = HomeworkData.Load(date);
            if (homeworkData.DueTimes.Count == 0)
            {
                foreach (string subject in currentSubjects)
                    homeworkData.DueTimes[subject] = appConfig.EveningClassCount >= 3 ? "晚修3" : "无";
                SaveHomeworkData();
            }
            if (EditMode) CreateTimeComboBoxes();
            if (_useWebView2) RefreshAllWebView2Content();
            else Invalidate();
        }

        private void SaveHomeworkData()
        {
            if (appConfig.MgmtEnabled && appConfig.MgmtForceRemote)
            {
                MessageBox.Show("当前已由管理端控制，无法保存本地修改。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            DateTime saveDate = historyMode && historyDate.HasValue ? historyDate.Value : currentDate;
            homeworkData.Save(saveDate);
            if (_useWebView2) RefreshAllWebView2Content();
            else Invalidate();
        }

        private void ToggleFullscreen()
        {
            fullscreen = !fullscreen;
            if (fullscreen)
            {
                FormBorderStyle = FormBorderStyle.None;
                WindowState = FormWindowState.Maximized;
            }
            else
            {
                FormBorderStyle = FormBorderStyle.Sizable;
                WindowState = FormWindowState.Normal;
                ClientSize = VIRTUAL_SIZE;
            }
            UpdateScale();
            if (_useWebView2) UpdateWebView2Positions();
            Invalidate();
        }

        private void UpdateScale()
        {
            Size clientSize = ClientSize;
            if (clientSize.Width == 0 || clientSize.Height == 0) return;
            float scaleW = (float)clientSize.Width / VIRTUAL_SIZE.Width;
            float scaleH = (float)clientSize.Height / VIRTUAL_SIZE.Height;
            scaleFactor = Math.Min(scaleW, scaleH);
            Size scaledVirtual = new Size((int)(VIRTUAL_SIZE.Width * scaleFactor), (int)(VIRTUAL_SIZE.Height * scaleFactor));
            offset = new Point((clientSize.Width - scaledVirtual.Width) / 2, (clientSize.Height - scaledVirtual.Height) / 2);
        }

        private Point MapToVirtual(Point screenPt) => new Point((int)((screenPt.X - offset.X) / scaleFactor), (int)((screenPt.Y - offset.Y) / scaleFactor));
        private Point MapToScreen(Point virtualPt) => new Point((int)(virtualPt.X * scaleFactor + offset.X), (int)(virtualPt.Y * scaleFactor + offset.Y));

        private Rectangle GetContentArea(int subjectIndex)
        {
            if (rotationMode)
            {
                var bigRect = new Rectangle(150, 100, 900, 475);
                return new Rectangle(bigRect.Left + 50, bigRect.Top + 150, bigRect.Width - 100, bigRect.Height - 200);
            }
            if (subjectIndex < 0 || subjectIndex >= gridRects.Count) return Rectangle.Empty;
            Rectangle rect = gridRects[subjectIndex];
            int topOffset = (appConfig.LastMode == "全科") ? 40 : 50;
            return new Rectangle(rect.Left + 10, rect.Top + topOffset + 10, rect.Width - 20, rect.Height - (topOffset + 20));
        }

        // ------------------------------ 下拉栏方法 ------------------------------
        private void InitializeDrawerPanel()
        {
            _drawerPanel = new BufferedPanel
            {
                BackColor = Color.FromArgb(200, 30, 30, 30),
                Location = new Point(0, 0),
                Size = new Size(ClientSize.Width, 0),
                Visible = true
            };
            _drawerPanel.Paint += DrawerPanel_Paint;
            _drawerPanel.MouseMove += (s, e) =>
            {
                Point screenPt = _drawerPanel.PointToScreen(e.Location);
                Point formPt = this.PointToClient(screenPt);
                OnMouseMove(new MouseEventArgs(e.Button, e.Clicks, formPt.X, formPt.Y, e.Delta));
            };
            this.Controls.Add(_drawerPanel);
            this.Controls.SetChildIndex(_drawerPanel, 0);
        }

        private void DrawerPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            DateTime now = DateTime.Now;
            string timeStr = now.ToString("HH:mm");
            string dateStr = now.ToString("yyyy年MM月dd日 dddd");
            using (Font timeFont = new Font("微软雅黑", 36, FontStyle.Bold))
            using (Font dateFont = new Font("微软雅黑", 16))
            using (Brush whiteBrush = new SolidBrush(Color.White))
            {
                SizeF timeSize = g.MeasureString(timeStr, timeFont);
                SizeF dateSize = g.MeasureString(dateStr, dateFont);
                float x = (_drawerPanel.Width - timeSize.Width) / 2;
                float y = (_drawerPanel.Height - (timeSize.Height + dateSize.Height + 10)) / 2;
                g.DrawString(timeStr, timeFont, whiteBrush, x, y);
                g.DrawString(dateStr, dateFont, whiteBrush, x, y + timeSize.Height + 10);
            }
        }

        private void OpenDrawer()
        {
            if (_drawerOpen) return;
            _drawerOpen = true;
            _suspendMouseGlow = true;
            int targetHeight = this.ClientSize.Height / MAX_DRAWER_HEIGHT_RATIO;
            SetDrawerHeight(targetHeight);
        }

        private void CloseDrawer()
        {
            if (!_drawerOpen) return;
            _drawerOpen = false;
            _suspendMouseGlow = false;
            AnimateDrawerClose();
        }

        private void SetDrawerHeight(int height)
        {
            if (_drawerPanel == null) return;
            _drawerPanel.Size = new Size(this.ClientSize.Width, height);
            _drawerPanel.Location = new Point(0, 0);
            _drawerPanel.Invalidate();
        }

        private void AnimateDrawerClose()
        {
            if (_drawerPanel == null) return;
            var anim = new Animation
            {
                Duration = TimeSpan.FromMilliseconds(200),
                Tag = "drawer"
            };
            int start = _drawerPanel.Height;
            anim.OnUpdate = (p) =>
            {
                int newHeight = (int)(start * (1 - p));
                SetDrawerHeight(newHeight);
            };
            anim.OnComplete = () => SetDrawerHeight(0);
            anim.Start();
            _animations.Add(anim);
        }

        // ------------------------------ 鼠标事件 ------------------------------
        private void OnMouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (!rotationMode && e.Y <= BAR_HEIGHT)
            {
                OpenDrawer();
            }
        }

        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            if (_isFlyingIn || _inRipple) return;

            Point v = MapToVirtual(e.Location);
            int x = v.X, y = v.Y;

            foreach (var key in _buttonPressed.Keys.ToList()) _buttonPressed[key] = false;
            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _minimizePressed = _closePressed = _expandPressed = _collapsePressed = _exportPressed = false;

            if (rotationMode)
            {
                if (backBtnRect.Contains(x, y)) { _backPressed = true; _buttonPressed["back"] = true; }
                if (editBtnRect.Contains(x, y)) { _editPressed = true; _buttonPressed["edit"] = true; }
                if (historyBtnRect.Contains(x, y)) { _historyPressed = true; _buttonPressed["history"] = true; }
                if (fullscreenBtnRect.Contains(x, y)) { _fullscreenPressed = true; _buttonPressed["fullscreen"] = true; }
                if (exportBtnRect.Contains(x, y)) { _exportPressed = true; _buttonPressed["export"] = true; }
            }
            else
            {
                if (!_buttonsExpanded)
                {
                    if (expandBtnRect.Contains(x, y)) { _expandPressed = true; _buttonPressed["expand"] = true; }
                    else if (editBtnRect.Contains(x, y)) { _editPressed = true; _buttonPressed["edit"] = true; }
                    else
                    {
                        bool isWin10 = Environment.OSVersion.Version.Major == 10 && Environment.OSVersion.Version.Minor == 0 && Environment.OSVersion.Version.Build < 22000;
                        if (isWin10 && minimizeBtnRect.Contains(x, y)) { _minimizePressed = true; _buttonPressed["minimize"] = true; }
                        else if (fullscreenBtnRect.Contains(x, y)) { _fullscreenPressed = true; _buttonPressed["fullscreen"] = true; }
                    }
                }
                else
                {
                    if (collapseBtnRect.Contains(x, y)) { _collapsePressed = true; _buttonPressed["collapse"] = true; }
                    else if (settingsBtnRect.Contains(x, y)) { _settingsPressed = true; _buttonPressed["settings"] = true; }
                    else if (rotateBtnRect.Contains(x, y)) { _rotatePressed = true; _buttonPressed["rotate"] = true; }
                    else if (exportBtnRect.Contains(x, y)) { _exportPressed = true; _buttonPressed["export"] = true; }
                    else if (editBtnRect.Contains(x, y)) { _editPressed = true; _buttonPressed["edit"] = true; }
                    else if (historyBtnRect.Contains(x, y)) { _historyPressed = true; _buttonPressed["history"] = true; }
                    else if (fullscreenBtnRect.Contains(x, y)) { _fullscreenPressed = true; _buttonPressed["fullscreen"] = true; }
                    else if (minimizeBtnRect.Contains(x, y)) { _minimizePressed = true; _buttonPressed["minimize"] = true; }
                    else if (closeBtnRect.Contains(x, y)) { _closePressed = true; _buttonPressed["close"] = true; }
                }
            }

            Invalidate();
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (_isFlyingIn || _inRipple) return;
            if (_suspendMouseGlow) return;

            if (!_isWin10)
            {
                if (!_mouseInside)
                {
                    _mouseInside = true;
                    if (_mouseGlowFadeOutAnimation != null && _mouseGlowFadeOutAnimation.IsRunning)
                    {
                        _animations.Remove(_mouseGlowFadeOutAnimation);
                        _mouseGlowFadeOutAnimation = null;
                    }
                    if (_mouseGlowAlpha <= 0.01f && _mouseGlowFadeAnimation == null)
                    {
                        var fadeIn = new Animation
                        {
                            Duration = TimeSpan.FromMilliseconds(500),
                            Tag = "mouseGlowFadeIn"
                        };
                        fadeIn.OnUpdate = (p) =>
                        {
                            _mouseGlowAlpha = (float)p;
                            Invalidate();
                        };
                        fadeIn.OnComplete = () =>
                        {
                            _mouseGlowAlpha = 1;
                            _mouseGlowFadeAnimation = null;
                            Invalidate();
                        };
                        fadeIn.Start();
                        _animations.Add(fadeIn);
                        _mouseGlowFadeAnimation = fadeIn;
                    }
                }

                _mouseVirtualPos = MapToVirtual(e.Location);
            }
            else
            {
                _mouseGlowAlpha = 0;
                _mouseVirtualPos = MapToVirtual(e.Location);
            }

            Invalidate();

            Point vBtn = _mouseVirtualPos;
            int x = vBtn.X, y = vBtn.Y;

            string newHoverKey = null;

            if (rotationMode)
            {
                if (backBtnRect.Contains(x, y)) newHoverKey = "back";
                else if (editBtnRect.Contains(x, y)) newHoverKey = "edit";
                else if (historyBtnRect.Contains(x, y)) newHoverKey = "history";
                else if (fullscreenBtnRect.Contains(x, y)) newHoverKey = "fullscreen";
                else if (exportBtnRect.Contains(x, y)) newHoverKey = "export";
            }
            else
            {
                if (!_buttonsExpanded)
                {
                    if (expandBtnRect.Contains(x, y)) newHoverKey = "expand";
                    else if (editBtnRect.Contains(x, y)) newHoverKey = "edit";
                    else if (fullscreenBtnRect.Contains(x, y)) newHoverKey = "fullscreen";
                }
                else
                {
                    if (collapseBtnRect.Contains(x, y)) newHoverKey = "collapse";
                    else if (settingsBtnRect.Contains(x, y)) newHoverKey = "settings";
                    else if (rotateBtnRect.Contains(x, y)) newHoverKey = "rotate";
                    else if (exportBtnRect.Contains(x, y)) newHoverKey = "export";
                    else if (editBtnRect.Contains(x, y)) newHoverKey = "edit";
                    else if (historyBtnRect.Contains(x, y)) newHoverKey = "history";
                    else if (fullscreenBtnRect.Contains(x, y)) newHoverKey = "fullscreen";
                    else if (minimizeBtnRect.Contains(x, y)) newHoverKey = "minimize";
                    else if (closeBtnRect.Contains(x, y)) newHoverKey = "close";
                }
            }

            if (newHoverKey == _currentHoverKey) return;

            if (_currentHoverKey != null)
                StartButtonHoverAnimation(_currentHoverKey, false);

            if (newHoverKey != null)
            {
                if (!_buttonHoverProgress.ContainsKey(newHoverKey))
                    _buttonHoverProgress[newHoverKey] = 0;
                StartButtonHoverAnimation(newHoverKey, true);
            }

            _currentHoverKey = newHoverKey;

            _hideMouseGlow = (inlineEditControl != null && inlineEditControl.Visible && inlineEditControl.Bounds.Contains(e.Location)) || _suspendMouseGlow;
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            if (_isFlyingIn || _inRipple) return;

            foreach (var key in _buttonPressed.Keys.ToList()) _buttonPressed[key] = false;
            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _minimizePressed = _closePressed = _expandPressed = _collapsePressed = _exportPressed = false;
            Invalidate();
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            if (_isFlyingIn) return;

            if (_currentHoverKey != null)
            {
                StartButtonHoverAnimation(_currentHoverKey, false);
                _currentHoverKey = null;
            }
            foreach (var key in _buttonPressed.Keys.ToList())
                _buttonPressed[key] = false;

            if (!_isWin10)
            {
                _mouseInside = false;
                if (_mouseGlowFadeAnimation != null && _mouseGlowFadeAnimation.IsRunning)
                    _animations.Remove(_mouseGlowFadeAnimation);
                if (_mouseGlowFadeOutAnimation != null && _mouseGlowFadeOutAnimation.IsRunning)
                    _animations.Remove(_mouseGlowFadeOutAnimation);

                float startAlpha = _mouseGlowAlpha;
                _mouseGlowFadeOutAnimation = new Animation
                {
                    Duration = TimeSpan.FromMilliseconds(500),
                    Tag = "mouseGlowFadeOut"
                };
                _mouseGlowFadeOutAnimation.OnUpdate = (p) =>
                {
                    _mouseGlowAlpha = startAlpha * (1 - (float)p);
                    Invalidate();
                };
                _mouseGlowFadeOutAnimation.OnComplete = () =>
                {
                    _mouseGlowAlpha = 0;
                    _mouseGlowFadeOutAnimation = null;
                    Invalidate();
                };
                _mouseGlowFadeOutAnimation.Start();
                _animations.Add(_mouseGlowFadeOutAnimation);
            }
            else
            {
                _mouseGlowAlpha = 0;
            }

            Invalidate();
        }

        private void OnActivated(object sender, EventArgs e) { UpdateScale(); Invalidate(); }

        private void StartButtonHoverAnimation(string key, bool fadeIn)
        {
            var existing = _animations.FirstOrDefault(a => a.Tag?.ToString() == key);
            if (existing != null) _animations.Remove(existing);

            float start = _buttonHoverProgress.GetValueOrDefault(key, 0);
            float target = fadeIn ? 1 : 0;

            var anim = new Animation
            {
                Duration = TimeSpan.FromMilliseconds(150),
                Tag = key
            };
            anim.OnUpdate = (p) =>
            {
                _buttonHoverProgress[key] = start + (target - start) * (float)p;
                Invalidate();
            };
            anim.OnComplete = () =>
            {
                _buttonHoverProgress[key] = target;
                Invalidate();
            };
            anim.Start();
            _animations.Add(anim);
        }

        // ------------------------------ 辅助绘图方法 ------------------------------
        private GraphicsPath CreateRoundedRectPath(Rectangle rect, int radius, bool bottomOnly = false)
        {
            var path = new GraphicsPath();
            if (radius <= 0) { path.AddRectangle(rect); return path; }

            if (bottomOnly)
            {
                int x = rect.X, y = rect.Y, w = rect.Width, h = rect.Height, d = radius * 2;
                path.AddLine(x, y, x + w, y);
                path.AddLine(x + w, y, x + w, y + h - radius);
                path.AddArc(new Rectangle(x + w - d, y + h - d, d, d), 0, 90);
                path.AddLine(x + w - radius, y + h, x + radius, y + h);
                path.AddArc(new Rectangle(x, y + h - d, d, d), 90, 90);
                path.AddLine(x, y + h - radius, x, y);
            }
            else
            {
                int d = radius * 2;
                Rectangle arcRect = new Rectangle(rect.Location, new Size(d, d));
                path.AddArc(arcRect, 180, 90);
                arcRect.X = rect.Right - d;
                path.AddArc(arcRect, 270, 90);
                arcRect.Y = rect.Bottom - d;
                path.AddArc(arcRect, 0, 90);
                arcRect.X = rect.Left;
                path.AddArc(arcRect, 90, 90);
                path.CloseFigure();
            }
            return path;
        }

        private GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            Rectangle arcRect = new Rectangle(rect.Location, new Size(d, d));
            path.AddArc(arcRect, 180, 90);
            arcRect.X = rect.Right - d;
            path.AddArc(arcRect, 270, 90);
            arcRect.Y = rect.Bottom - d;
            path.AddArc(arcRect, 0, 90);
            arcRect.X = rect.Left;
            path.AddArc(arcRect, 90, 90);
            path.CloseFigure();
            return path;
        }

        private Rectangle GetDueTimeRect(int subjectIndex)
        {
            Rectangle rect = gridRects[subjectIndex];
            string subject = currentSubjects[subjectIndex];
            string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
            string prefix = "提交时间：";
            string displayTime = string.IsNullOrEmpty(dueTime) ? "无" : dueTime;

            float prefixWidth, timeWidth;
            using (Graphics g = CreateGraphics())
            {
                prefixWidth = g.MeasureString(prefix, fontSmall).Width;
                timeWidth = g.MeasureString(displayTime, font22).Width;
            }
            float totalWidth = prefixWidth + timeWidth;
            int rightX = rect.Right - 10;
            float startX = rightX - totalWidth;
            int lineY = rect.Top + 50;
            int timeY = lineY - font22.Height;
            int prefixY = lineY - fontSmall.Height;
            int minY = Math.Min(prefixY, timeY);
            int maxY = lineY;
            int height = maxY - minY + 2;
            if (height < 25) { minY = maxY - 25; height = 25; }
            return new Rectangle((int)startX, minY, (int)totalWidth + 5, height);
        }

        // ------------------------------ 模式切换下拉框 ------------------------------
        private void InitializeModeComboBox()
        {
            modeComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 10),
                Width = 100,
                Height = 25,
                Visible = false,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(40, 40, 40),
                ForeColor = Color.White
            };
            modeComboBox.Items.AddRange(new object[] { "大理", "中理", "小理", "大文", "全科" });
            modeComboBox.SelectedItem = appConfig.LastMode;
            modeComboBox.SelectedIndexChanged += ModeComboBox_SelectedIndexChanged;
            modeComboBox.DrawMode = DrawMode.OwnerDrawFixed;
            modeComboBox.DrawItem += ModeComboBox_DrawItem;
            Controls.Add(modeComboBox);
        }

        private void ModeComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var combo = sender as ComboBox;
            string text = combo.Items[e.Index].ToString();
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bgColor = selected ? Color.FromArgb(80, 80, 80) : Color.FromArgb(40, 40, 40);
            using (var bgBrush = new SolidBrush(bgColor))
                e.Graphics.FillRectangle(bgBrush, e.Bounds);
            using (var textBrush = new SolidBrush(Color.White))
                e.Graphics.DrawString(text, combo.Font, textBrush, e.Bounds.Left + 5, e.Bounds.Top + 2);
            e.DrawFocusRectangle();
        }

        private void ModeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string newMode = modeComboBox.SelectedItem.ToString();
            if (newMode == appConfig.LastMode) return;
            appConfig.LastMode = newMode;
            appConfig.Save();
            UpdateSubjectsByMode();

            // 强制重建所有控件
            bool wasEditMode = EditMode;
            bool wasRotationMode = rotationMode;
            if (wasRotationMode)
            {
                rotationMode = false;
                rotationTimer.Stop();
                if (rotationWebView != null)
                {
                    rotationWebView.Dispose();
                    rotationWebView = null;
                }
            }
            if (wasEditMode)
                EditMode = false;
            else
            {
                if (_useWebView2) DestroyAllWebView2();
                DestroyTimeComboBoxes();
            }

            CalculateGridRects();

            if (wasRotationMode)
            {
                rotationMode = true;
                rotationIndex = 0;
                rotationTimer.Start();
                if (_useWebView2) CreateRotationWebView();
                Invalidate();
            }
            else if (wasEditMode)
                EditMode = true;
            else
            {
                if (_useWebView2)
                    RecreateWebView2ForGrid();
                else
                    Invalidate(); // GDI+ 模式重绘即可
            }
            Invalidate();
        }

        // ------------------------------ 鼠标点击与导出 ------------------------------
        private void OnMouseClick(object sender, MouseEventArgs e)
        {
            Point v = MapToVirtual(e.Location);
            int x = v.X, y = v.Y;

            if (_drawerOpen && _drawerPanel != null && !_drawerPanel.Bounds.Contains(e.Location))
                CloseDrawer();

            // 涟漪动画（Win10 下不启用）
            if (!_isWin10)
            {
                _rippleCenter = v;
                _rippleRadius = 0;
                _rippleAlpha = 180;
                float maxRadius = (float)Math.Sqrt(VIRTUAL_SIZE.Width * VIRTUAL_SIZE.Width + VIRTUAL_SIZE.Height * VIRTUAL_SIZE.Height) * 1.2f;

                if (_mouseGlowFadeAnimation != null) _animations.Remove(_mouseGlowFadeAnimation);
                _mouseGlowAlpha = 0;
                if (_rippleAnimation != null) _animations.Remove(_rippleAnimation);

                _inRipple = true;

                _rippleAnimation = new Animation
                {
                    Duration = TimeSpan.FromMilliseconds(500),
                    IsLooping = false,
                    Tag = "ripple"
                };
                _rippleAnimation.OnUpdate = (p) =>
                {
                    _rippleRadius = maxRadius * (float)p;
                    _rippleAlpha = (int)(180 * (1 - p));
                    Invalidate();
                };
                _rippleAnimation.OnComplete = () =>
                {
                    _rippleCenter = new Point(-1, -1);
                    _rippleRadius = 0;
                    _rippleAlpha = 0;
                    _inRipple = false;
                    Invalidate();

                    var fadeIn = new Animation
                    {
                        Duration = TimeSpan.FromMilliseconds(500),
                        Tag = "mouseGlowFadeIn"
                    };
                    fadeIn.OnUpdate = (p) =>
                    {
                        _mouseGlowAlpha = (float)p;
                        Invalidate();
                    };
                    fadeIn.OnComplete = () =>
                    {
                        _mouseGlowAlpha = 1;
                        _mouseGlowFadeAnimation = null;
                        Invalidate();
                    };
                    fadeIn.Start();
                    _animations.Add(fadeIn);
                    _mouseGlowFadeAnimation = fadeIn;
                };
                _rippleAnimation.Start();
                _animations.Add(_rippleAnimation);
            }
            else
            {
                _mouseGlowAlpha = 0;
            }

            // 轮播模式
            if (rotationMode)
            {
                if (backBtnRect.Contains(x, y))
                {
                    rotationMode = false;
                    rotationTimer.Stop();
                    if (rotationWebView != null)
                    {
                        rotationWebView.Dispose();
                        rotationWebView = null;
                    }
                    if (EditMode)
                    {
                        CreatePlainTextEditors();
                        CreateTimeComboBoxes();
                    }
                    else
                    {
                        if (_useWebView2)
                            RecreateWebView2ForGrid();
                        // GDI+ 模式无需重建，只需重绘
                    }
                    StartCardFlyInAnimation();
                    Invalidate();
                }
                else if (fullscreenBtnRect.Contains(x, y))
                {
                    if (EditMode) EditMode = false;
                    ToggleFullscreen();
                }
                return;
            }

            // 非轮播模式
            bool isWin10 = Environment.OSVersion.Version.Major == 10 && Environment.OSVersion.Version.Minor == 0 && Environment.OSVersion.Version.Build < 22000;

            if (!_buttonsExpanded)
            {
                if (expandBtnRect.Contains(x, y))
                {
                    _buttonsExpanded = true;
                    Invalidate();
                    return;
                }
                else if (editBtnRect.Contains(x, y))
                {
                    if (EditMode) SaveHomeworkData();
                    EditMode = !EditMode;
                    return;
                }
                else
                {
                    if (isWin10 && minimizeBtnRect.Contains(x, y))
                    {
                        WindowState = FormWindowState.Minimized;
                        return;
                    }
                    else if (fullscreenBtnRect.Contains(x, y))
                    {
                        if (EditMode) EditMode = false;
                        ToggleFullscreen();
                        return;
                    }
                }
            }
            else
            {
                if (collapseBtnRect.Contains(x, y))
                {
                    _buttonsExpanded = false;
                    Invalidate();
                    return;
                }
                else if (settingsBtnRect.Contains(x, y))
                {
                    using (var settingsForm = new SettingsForm(this))
                    {
                        settingsForm.ShowDialog();
                    }
                    return;
                }
                else if (rotateBtnRect.Contains(x, y))
                {
                    bool hasContent = currentSubjects.Any(subj => homeworkData.Subjects.ContainsKey(subj) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[subj]));
                    if (!hasContent)
                    {
                        MessageBox.Show("所有科目都没有作业内容，无法进入轮播模式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (EditMode) EditMode = false;
                    if (_useWebView2)
                        DestroyAllWebView2();
                    DestroyTimeComboBoxes();
                    rotationMode = true;
                    rotationIndex = 0;
                    rotationTimer.Start();
                    if (_useWebView2) CreateRotationWebView();
                    Invalidate();
                    return;
                }
                else if (exportBtnRect.Contains(x, y))
                {
                    ShowExportDialog();
                    return;
                }
                else if (editBtnRect.Contains(x, y))
                {
                    if (EditMode) SaveHomeworkData();
                    EditMode = !EditMode;
                    return;
                }
                else if (historyBtnRect.Contains(x, y))
                {
                    if (historyMode)
                    {
                        historyMode = false;
                        historyDate = null;
                        LoadHomeworkData(currentDate);
                        Invalidate();
                    }
                    else
                    {
                        using (var dlg = new HistoryDialog())
                        {
                            if (dlg.ShowDialog(this) == DialogResult.OK && dlg.SelectedDate.HasValue)
                            {
                                historyDate = dlg.SelectedDate.Value;
                                LoadHomeworkData(historyDate.Value);
                                historyMode = true;
                                Invalidate();
                            }
                        }
                    }
                    return;
                }
                else if (minimizeBtnRect.Contains(x, y))
                {
                    WindowState = FormWindowState.Minimized;
                    return;
                }
                else if (closeBtnRect.Contains(x, y))
                {
                    Application.Exit();
                    return;
                }
                else if (fullscreenBtnRect.Contains(x, y))
                {
                    if (EditMode) EditMode = false;
                    ToggleFullscreen();
                    return;
                }
            }
        }

        // ------------------------------ 轮播切换（无动画） ------------------------------
        private void RotateManual(int direction)
        {
            var nonEmpty = new List<int>();
            for (int i = 0; i < currentSubjects.Length; i++)
                if (homeworkData.Subjects.ContainsKey(currentSubjects[i]) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[currentSubjects[i]]))
                    nonEmpty.Add(i);
            if (nonEmpty.Count == 0)
            {
                rotationMode = false;
                rotationTimer.Stop();
                StartCardFlyInAnimation();
                MessageBox.Show("所有科目都没有作业内容，已退出轮播模式", "提示");
                return;
            }

            int curIdx = nonEmpty.IndexOf(rotationIndex);
            if (curIdx < 0) curIdx = 0;
            int newIdx = (curIdx + direction + nonEmpty.Count) % nonEmpty.Count;
            int newIndex = nonEmpty[newIdx];

            if (newIndex == rotationIndex) return;

            rotationIndex = newIndex;
            if (rotationWebView != null)
            {
                string subject = currentSubjects[rotationIndex];
                string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
                _ = UpdateWebView2Content(rotationWebView, content);
            }
            else
            {
                // GDI+ 模式下的轮播（虽然当前轮播只在 WebView2 模式下支持，但保留逻辑）
                Invalidate();
            }
            Invalidate();
        }

        private void RotateNext() => RotateManual(1);

        // ------------------------------ 图片加载 ------------------------------
        private void LoadImages()
        {
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
            buttonImage = LoadImage(Path.Combine(imagePath, "按钮.png"), new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));
            historyBtnImage = LoadImage(Path.Combine(imagePath, "更多.png"), new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));

            Image originalBack = LoadImage(Path.Combine(imagePath, "返回.png"), true);
            if (originalBack != null)
            {
                int targetHeight = BTN_SQUARE_SIZE;
                int targetWidth = (int)((float)originalBack.Width / originalBack.Height * targetHeight);
                backBtnImage = new Bitmap(originalBack, new Size(targetWidth, targetHeight));
            }

            string arrowPath = Path.Combine(imagePath, "箭头图片.png");
            if (File.Exists(arrowPath))
            {
                leftArrowImage = Image.FromFile(arrowPath);
                rightArrowImage = (Image)new Bitmap(leftArrowImage);
                rightArrowImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            }

            string[] iconNames = { "编辑", "返回", "历史", "轮换", "全屏", "设置", "缩小", "完成", "关闭", "最小化", "展开", "收起", "导出" };
            int iconSize = (int)(BTN_SQUARE_SIZE * 0.6);
            foreach (string name in iconNames)
            {
                string filePath = Path.Combine(imagePath, name + ".png");
                if (File.Exists(filePath))
                {
                    using (var img = Image.FromFile(filePath))
                    {
                        var targetBitmap = new Bitmap(iconSize, iconSize);
                        using (var g = Graphics.FromImage(targetBitmap))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(img, 0, 0, iconSize, iconSize);
                        }
                        buttonIcons[name] = targetBitmap;
                    }
                }
            }
        }

        private Image LoadImage(string path, bool originalSize) { try { return File.Exists(path) ? Image.FromFile(path) : null; } catch { return null; } }
        private Image LoadImage(string path, Size size) { try { return File.Exists(path) ? new Bitmap(Image.FromFile(path), size) : null; } catch { return null; } }

        // ------------------------------ 颜色辅助 ------------------------------
        private Color ParseColor(string rgbString, Color defaultColor)
        {
            try
            {
                var parts = rgbString.Split(',');
                if (parts.Length == 3) return Color.FromArgb(int.Parse(parts[0]), int.Parse(parts[1]), int.Parse(parts[2]));
            }
            catch { }
            return defaultColor;
        }

        private Brush GetDueTimeBrush(string value)
        {
            if (value.StartsWith("晚修") && int.TryParse(value.Substring(2), out int idx))
            {
                idx--;
                return idx switch
                {
                    0 => RED_SEMI,
                    1 => ORANGE_SEMI,
                    2 => GREEN_SEMI,
                    3 => BLUE_SEMI,
                    4 => PURPLE_SEMI,
                    5 => DARKORANGE_SEMI,
                    _ => new SolidBrush(TEXT_COLOR)
                };
            }
            return new SolidBrush(TEXT_COLOR);
        }

        // ------------------------------ 主绘制 ------------------------------
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            g.TranslateTransform(offset.X, offset.Y);
            g.ScaleTransform(scaleFactor, scaleFactor);

            if (_inSizeMove)
            {
                using (var bgBrush = new SolidBrush(Color.FromArgb(0, 0, 0, 0)))
                {
                    g.FillRectangle(bgBrush, new Rectangle(0, 0, VIRTUAL_SIZE.Width, VIRTUAL_SIZE.Height));
                }
                g.ResetTransform();
                return;
            }

            // 背景绘制
            if (appConfig.UseBackgroundImage && _backgroundImage != null)
                g.DrawImage(_backgroundImage, new Rectangle(0, 0, VIRTUAL_SIZE.Width, VIRTUAL_SIZE.Height));
            else if (!_micaEnabled)
            {
                using (var bgBrush = new SolidBrush(Color.FromArgb(_backgroundAlpha, 32, 32, 32)))
                    g.FillRectangle(bgBrush, new Rectangle(0, 0, VIRTUAL_SIZE.Width, VIRTUAL_SIZE.Height));
            }

            // 鼠标跟随光晕
            if (!_isWin10 && appConfig.ShowMouseGlow && !_hideMouseGlow && _mouseVirtualPos.X > 0 && _mouseVirtualPos.X < VIRTUAL_SIZE.Width &&
                _mouseVirtualPos.Y > 0 && _mouseVirtualPos.Y < VIRTUAL_SIZE.Height && _mouseGlowAlpha > 0)
            {
                int centerX = _mouseVirtualPos.X;
                int centerY = _mouseVirtualPos.Y;
                int radius = 120;
                if (radius > 0)
                {
                    using (var path = new GraphicsPath())
                    {
                        path.AddEllipse(centerX - radius, centerY - radius, radius * 2, radius * 2);
                        if (path.PointCount > 0)
                        {
                            using (var brush = new PathGradientBrush(path))
                            {
                                int alpha = (int)(80 * _mouseGlowAlpha);
                                brush.CenterColor = Color.FromArgb(alpha, 255, 255, 255);
                                brush.SurroundColors = new Color[] { Color.FromArgb(0, 255, 255, 255) };
                                brush.CenterPoint = new PointF(centerX, centerY);
                                g.FillPath(brush, path);
                            }
                        }
                    }
                }
            }

            // 顶部条
            Rectangle barRect = new Rectangle(0, BAR_Y, VIRTUAL_SIZE.Width, BAR_HEIGHT);
            if (gridRects.Count >= 3)
            {
                int barLeft = gridRects[0].Left;
                int barRight = gridRects[2].Right;
                barRect = new Rectangle(barLeft, BAR_Y, barRight - barLeft, BAR_HEIGHT);
            }

            float opacityFactor = appConfig.CardOpacity / 100f;
            int barAlpha = (int)(255 * opacityFactor);
            Color barColor = ParseColor(appConfig.BarColor, Color.Yellow);
            using (var barBrush = new SolidBrush(Color.FromArgb(barAlpha, barColor.R, barColor.G, barColor.B)))
            using (var barPath = CreateRoundedRectPath(barRect, ROUND_RADIUS, bottomOnly: true))
            {
                g.FillPath(barBrush, barPath);
                using (var borderPen = new Pen(Color.FromArgb(120, 255, 255, 255), 1))
                    g.DrawPath(borderPen, barPath);
            }

            // 日期信息
            DrawDateInfo(g, barRect);

            // 网格或轮播视图
            if (rotationMode)
                DrawRotationViewBackground(g);
            else
                DrawGridBackground(g);

            // 按钮栏
            DrawButtons(g, barRect);

            // 涟漪效果
            if (!_isWin10 && _rippleCenter.X >= 0 && _rippleCenter.Y >= 0 && _rippleRadius > 0 && _rippleAlpha > 0)
            {
                float radius = _rippleRadius;
                using (var path = new GraphicsPath())
                {
                    path.AddEllipse(_rippleCenter.X - radius, _rippleCenter.Y - radius, radius * 2, radius * 2);
                    if (path.PointCount > 0)
                    {
                        using (var brush = new PathGradientBrush(path))
                        {
                            brush.CenterColor = Color.FromArgb((int)_rippleAlpha, 255, 255, 255);
                            brush.SurroundColors = new Color[] { Color.FromArgb(0, 255, 255, 255) };
                            brush.CenterPoint = new PointF(_rippleCenter.X, _rippleCenter.Y);
                            g.FillPath(brush, path);
                        }
                    }
                }
            }

            g.ResetTransform();
        }

        private void DrawDateInfo(Graphics g, Rectangle barRect)
        {
            DateTime now = historyMode && historyDate.HasValue ? historyDate.Value : currentDate;
            string[] weekdays = { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日" };
            string weekday = weekdays[(int)now.DayOfWeek == 0 ? 6 : (int)now.DayOfWeek - 1];
            string date = now.ToString("yyyy年MM月dd日") + (historyMode ? " (历史作业)" : "");

            using (Brush white = new SolidBrush(Color.White))
            {
                if (rotationMode)
                {
                    int totalHeight = font20.Height + font12.Height;
                    int y = barRect.Top + (barRect.Height - totalHeight) / 2;
                    SizeF wSize = g.MeasureString(weekday, font20);
                    g.DrawString(weekday, font20, white, (VIRTUAL_SIZE.Width - wSize.Width) / 2, y);
                    SizeF dSize = g.MeasureString(date, font12);
                    g.DrawString(date, font12, white, (VIRTUAL_SIZE.Width - dSize.Width) / 2, y + font20.Height);
                }
                else
                {
                    int totalHeight = font20.Height + font12.Height;
                    int y = barRect.Top + (barRect.Height - totalHeight) / 2;
                    int x = barRect.Left + BAR_PADDING;
                    g.DrawString(weekday, font20, white, x, y);
                    g.DrawString(date, font12, white, x, y + font20.Height);
                }
            }
        }

        // ------------------------------ 网格背景绘制（GDI+ 模式同时绘制文本并支持滚动） ------------------------------
        private void DrawGridBackground(Graphics g)
        {
            float opacityFactor = appConfig.CardOpacity / 100f;
            for (int i = 0; i < gridRects.Count; i++)
            {
                Rectangle rect = gridRects[i];
                float offsetX = 0;
                int alpha = 255;
                if (_cardAnimations.TryGetValue(i, out var anim) && anim.IsRunning)
                {
                    double eased = 1 - Math.Pow(1 - anim.Progress, 2);
                    float startX = VIRTUAL_SIZE.Width;
                    float targetX = rect.X;
                    offsetX = (startX - targetX) * (1 - (float)eased);
                    alpha = (int)(255 * eased);
                }
                Rectangle drawRect = new Rectangle(rect.X + (int)offsetX, rect.Y, rect.Width, rect.Height);

                using (var shadowPath = CreateRoundedRectPath(new Rectangle(drawRect.X + 3, drawRect.Y + 3, drawRect.Width, drawRect.Height), ROUND_RADIUS))
                using (var shadowBrush = new SolidBrush(Color.FromArgb(30, 0, 0, 0)))
                    g.FillPath(shadowBrush, shadowPath);

                int topAlpha = (int)(220 * opacityFactor * alpha / 255.0);
                int bottomAlpha = (int)(160 * opacityFactor * alpha / 255.0);
                using (var path = CreateRoundedRectPath(drawRect, ROUND_RADIUS))
                using (var bgBrush = new LinearGradientBrush(drawRect, Color.FromArgb(topAlpha, 255, 255, 255), Color.FromArgb(bottomAlpha, 240, 240, 255), LinearGradientMode.Vertical))
                {
                    g.FillPath(bgBrush, path);
                    string subject = currentSubjects[i];
                    string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
                    bool highlightActive = _activeEvenings.Contains(dueTime);
                    bool highlightFlash = _flashingEvenings.Contains(dueTime);
                    Color borderColor = Color.FromArgb(100, 128, 128, 128);
                    float borderWidth = 1;
                    if (_debugFlashing || highlightFlash)
                    {
                        using (var redBrush = new SolidBrush(Color.FromArgb(flashStep, 255, 0, 0)))
                            g.FillPath(redBrush, path);
                        DrawLaserBorder(g, path, drawRect);
                        borderColor = Color.FromArgb(200, 255, 0, 0);
                    }
                    else if (highlightActive) { borderColor = SystemColors.Highlight; borderWidth = 2; }
                    using (var pen = new Pen(borderColor, borderWidth))
                        g.DrawPath(pen, path);
                }

                // 绘制科目名和提交时间
                if (i < currentSubjects.Length)
                {
                    string subject = currentSubjects[i];
                    Font subjectFont = font22;
                    Font dueTimeFont = font22;
                    Font dueTimeLabelFont = fontSmall;
                    int topOffset = 50;
                    if (appConfig.LastMode == "全科")
                    {
                        subjectFont = font20;
                        dueTimeFont = font20;
                        dueTimeLabelFont = fontSmall;
                        topOffset = 40;
                    }
                    int subjectY = drawRect.Top + topOffset - subjectFont.Height;
                    g.DrawString(subject, subjectFont, new SolidBrush(TEXT_COLOR), drawRect.Left + 10, subjectY);
                    int lineY = drawRect.Top + topOffset;
                    g.DrawLine(Pens.Gray, drawRect.Left + 10, lineY, drawRect.Right - 10, lineY);

                    if (appConfig.ShowDueTime && !EditMode)
                    {
                        string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
                        string prefix = "提交时间：";
                        string displayTime = string.IsNullOrEmpty(dueTime) ? "无" : dueTime;
                        float prefixWidth = g.MeasureString(prefix, dueTimeLabelFont).Width;
                        float timeWidth = g.MeasureString(displayTime, dueTimeFont).Width;
                        float totalWidth = prefixWidth + timeWidth;
                        int rightX = drawRect.Right - 10;
                        float startX = rightX - totalWidth;
                        int timeY = lineY - dueTimeFont.Height;
                        int prefixY = lineY - dueTimeLabelFont.Height;
                        g.DrawString(prefix, dueTimeLabelFont, new SolidBrush(TEXT_COLOR), startX, prefixY);
                        Brush timeBrush = GetDueTimeBrush(displayTime);
                        g.DrawString(displayTime, dueTimeFont, timeBrush, startX + prefixWidth, timeY);
                        if (timeBrush != RED_SEMI && timeBrush != ORANGE_SEMI && timeBrush != GREEN_SEMI && timeBrush != BLUE_SEMI && timeBrush != PURPLE_SEMI && timeBrush != DARKORANGE_SEMI) timeBrush.Dispose();
                    }

                    // 绘制作业内容（GDI+ 模式，支持滚动）
                    if (!editMode && !_useWebView2)
                    {
                        Rectangle textArea = new Rectangle(drawRect.Left + 10, drawRect.Top + topOffset + 10, drawRect.Width - 20, drawRect.Height - (topOffset + 20));
                        string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
                        if (!string.IsNullOrWhiteSpace(content))
                        {
                            float offset = scrollOffsets.ContainsKey(i) ? scrollOffsets[i] : 0f;
                            DrawTextInArea(g, content, textArea, font30, false, offset);
                        }
                        else
                        {
                            using (var textPath = CreateRoundedRectPath(textArea, ROUND_RADIUS))
                            using (var lightBrush = new SolidBrush(Color.FromArgb((int)(255 * opacityFactor), 255, 255, 255)))
                                g.FillPath(lightBrush, textPath);
                            string hint = editMode ? "点我编辑作业" : "今天暂时没有此项作业";
                            g.DrawString(hint, hintFont, RED_SEMI, textArea, CenterStringFormat());
                        }
                    }
                }
            }

            // WebView2 模式下的控件创建（仅在非编辑非轮播时）
            if (!rotationMode && !editMode && _useWebView2 && webView2Initialized)
            {
                for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
                {
                    if (i >= cardWebViews.Count)
                    {
                        Rectangle textArea = GetContentArea(i);
                        CreateWebView2ForCard(i, textArea);
                    }
                    else
                    {
                        Rectangle textArea = GetContentArea(i);
                        UpdateWebView2Position(cardWebViews[i], textArea);
                    }
                }
                for (int i = currentSubjects.Length; i < cardWebViews.Count; i++)
                    cardWebViews[i].Visible = false;
            }
        }

        // ------------------------------ 轮播模式背景绘制（仅背景，WebView 单独控制） ------------------------------
        private void DrawRotationViewBackground(Graphics g)
        {
            float opacityFactor = appConfig.CardOpacity / 100f;
            int topAlpha = (int)(220 * opacityFactor);
            int bottomAlpha = (int)(160 * opacityFactor);
            Rectangle cardRect = new Rectangle(150, 100, 900, 475);
            DrawSingleCardBackground(g, cardRect, topAlpha, bottomAlpha);

            if (rotationMode && !editMode && _useWebView2 && webView2Initialized)
            {
                if (rotationWebView == null)
                    CreateRotationWebView();
                else
                {
                    Rectangle webViewRect = new Rectangle(150, 230, 900, 345);
                    UpdateWebView2Position(rotationWebView, webViewRect);
                }
            }
        }

        private void DrawSingleCardBackground(Graphics g, Rectangle rect, int topAlpha, int bottomAlpha)
        {
            using (var path = CreateRoundedRectPath(rect, ROUND_RADIUS))
            using (var bgBrush = new LinearGradientBrush(rect, Color.FromArgb(topAlpha, 255, 255, 255), Color.FromArgb(bottomAlpha, 240, 240, 255), LinearGradientMode.Vertical))
            {
                g.FillPath(bgBrush, path);
                using (var pen = new Pen(Color.FromArgb(100, 128, 128, 128), 1))
                    g.DrawPath(pen, path);
            }

            string subject = rotationIndex < currentSubjects.Length ? currentSubjects[rotationIndex] : "";
            SizeF subjectSize = g.MeasureString(subject, font36);
            g.DrawString(subject, font36, Brushes.White, rect.X + (rect.Width - subjectSize.Width) / 2, rect.Y + 30);

            int lineY = rect.Y + 120;
            g.DrawLine(new Pen(Color.Gray, 2), rect.X + 50, lineY, rect.Right - 50, lineY);

            if (appConfig.ShowDueTime)
            {
                string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
                string displayTime = string.IsNullOrEmpty(dueTime) ? "无" : dueTime;
                string prefix = "提交时间：";
                float prefixWidth = g.MeasureString(prefix, fontSmall).Width;
                float timeWidth = g.MeasureString(displayTime, font22).Width;
                float totalWidth = prefixWidth + timeWidth;
                int rightX = rect.Right - 60;
                float startX = rightX - totalWidth;
                int prefixY = lineY - fontSmall.Height;
                int timeY = lineY - font22.Height;
                g.DrawString(prefix, fontSmall, new SolidBrush(TEXT_COLOR), startX, prefixY);
                Brush timeBrush = GetDueTimeBrush(displayTime);
                g.DrawString(displayTime, font22, timeBrush, startX + prefixWidth, timeY);
                if (timeBrush != RED_SEMI && timeBrush != ORANGE_SEMI && timeBrush != GREEN_SEMI && timeBrush != BLUE_SEMI && timeBrush != PURPLE_SEMI && timeBrush != DARKORANGE_SEMI) timeBrush.Dispose();
            }
        }

        // ------------------------------ GDI+ 文本绘制（支持滚动） ------------------------------
        private void DrawTextInArea(Graphics g, string text, Rectangle area, Font font, bool center = false, float scrollOffset = 0f)
        {
            var lines = new List<string>();
            foreach (string para in text.Split('\n'))
            {
                if (string.IsNullOrEmpty(para)) continue;
                string[] words = para.Split(' ');
                string line = "";
                foreach (string word in words)
                {
                    string test = line + (line == "" ? "" : " ") + word;
                    if (g.MeasureString(test, font).Width <= area.Width) line = test;
                    else
                    {
                        if (!string.IsNullOrEmpty(line)) lines.Add(line);
                        if (g.MeasureString(word, font).Width > area.Width)
                        {
                            string remaining = word;
                            while (remaining.Length > 0)
                            {
                                int take = 1;
                                while (take < remaining.Length && g.MeasureString(remaining.Substring(0, take), font).Width <= area.Width) take++;
                                while (take > 0 && g.MeasureString(remaining.Substring(0, take), font).Width > area.Width) take--;
                                if (take == 0) take = 1;
                                lines.Add(remaining.Substring(0, take));
                                remaining = remaining.Substring(take);
                            }
                            line = "";
                        }
                        else line = word;
                    }
                }
                if (!string.IsNullOrEmpty(line)) lines.Add(line);
            }

            float lineHeight = font.GetHeight(g);
            var state = g.Save();
            g.SetClip(area);
            using (var brush = new SolidBrush(TEXT_COLOR))
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    float y = area.Top + i * lineHeight - scrollOffset;
                    if (center)
                    {
                        SizeF sz = g.MeasureString(lines[i], font);
                        g.DrawString(lines[i], font, brush, area.Left + (area.Width - sz.Width) / 2, y);
                    }
                    else g.DrawString(lines[i], font, brush, area.Left + 5, y);
                }
            }
            g.Restore(state);
        }

        private float MeasureTextHeight(string text, Font font, int maxWidth)
        {
            using (var g = CreateGraphics())
                return g.MeasureString(text, font, maxWidth).Height;
        }

        private void DrawLaserBorder(Graphics g, GraphicsPath path, Rectangle rect)
        {
            using (var pen = new Pen(Color.Red, 3) { DashStyle = DashStyle.Custom, DashPattern = new float[] { 8, 8 }, DashOffset = _laserOffset * 16 })
                g.DrawPath(pen, path);
        }

        private StringFormat CenterStringFormat() => new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

        // ------------------------------ 滚动定时器（仅 GDI+ 模式） ------------------------------
        private void ScrollTimer_Tick(object sender, EventArgs e)
        {
            if (editMode) return;
            if (_useWebView2) return; // WebView2 模式不处理滚动

            bool needRedraw = false;
            float speed = appConfig.ScrollSpeed * 0.05f;
            float epsilon = 0.01f;
            for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
            {
                string subject = currentSubjects[i];
                string content = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
                if (string.IsNullOrWhiteSpace(content)) continue;

                Rectangle textArea = GetContentArea(i);
                float contentHeight = MeasureTextHeight(content, font30, textArea.Width);
                if (contentHeight <= textArea.Height) continue;

                if (!scrollOffsets.ContainsKey(i)) scrollOffsets[i] = 0f;
                if (!scrollPaused.ContainsKey(i)) scrollPaused[i] = false;
                if (!pauseStartTime.ContainsKey(i)) pauseStartTime[i] = DateTime.Now;
                if (speed <= 0) continue;

                if (scrollPaused[i])
                {
                    if ((DateTime.Now - pauseStartTime[i]).TotalSeconds >= SCROLL_PAUSE_SECONDS)
                    {
                        if (scrollOffsets[i] >= contentHeight - textArea.Height - epsilon)
                        {
                            scrollOffsets[i] = 0f;
                            pauseStartTime[i] = DateTime.Now;
                        }
                        else { scrollPaused[i] = false; scrollOffsets[i] = speed; }
                    }
                }
                else
                {
                    if (scrollOffsets[i] <= epsilon) { scrollPaused[i] = true; pauseStartTime[i] = DateTime.Now; }
                    else
                    {
                        scrollOffsets[i] += speed;
                        if (scrollOffsets[i] >= contentHeight - textArea.Height - epsilon)
                        {
                            scrollOffsets[i] = contentHeight - textArea.Height;
                            scrollPaused[i] = true;
                            pauseStartTime[i] = DateTime.Now;
                        }
                    }
                }
                needRedraw = true;
            }
            if (needRedraw) Invalidate();
        }

        // ------------------------------ 按钮绘制 ------------------------------
        private void DrawButtons(Graphics g, Rectangle barRect)
        {
            if (rotationMode)
            {
                int backY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;
                backBtnRect = new Rectangle(35, backY, BTN_SQUARE_SIZE, BTN_SQUARE_SIZE);
                DrawWin11Button(g, backBtnRect, "返回", buttonIcons.ContainsKey("返回") ? buttonIcons["返回"] : null, ref backBtnRect, backBtnRect.Location, _backPressed, "back");

                int btnY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;
                int totalButtonsWidth = 3 * BTN_SQUARE_SIZE + 2 * 10;
                int startX = barRect.Right - totalButtonsWidth - 20;

                Point editPos = new Point(startX, btnY);
                Point historyPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point fullscreenPos = new Point(historyPos.X + BTN_SQUARE_SIZE + 10, btnY);

                string editText = editMode ? "完成" : "编辑";
                string editIcon = editMode ? "完成" : "编辑";
                DrawWin11Button(g, editBtnRect, editText, buttonIcons.ContainsKey(editIcon) ? buttonIcons[editIcon] : null, ref editBtnRect, editPos, _editPressed, "edit");

                string historyIcon = historyMode ? "返回" : "历史";
                DrawWin11Button(g, historyBtnRect, historyMode ? "返回" : "历史", buttonIcons.ContainsKey(historyIcon) ? buttonIcons[historyIcon] : null, ref historyBtnRect, historyPos, _historyPressed, "history");

                string fullscreenIcon = fullscreen ? "缩小" : "全屏";
                DrawWin11Button(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreenIcon) ? buttonIcons[fullscreenIcon] : null, ref fullscreenBtnRect, fullscreenPos, _fullscreenPressed, "fullscreen");

                // 左右箭头已移除
            }
            else
            {
                bool isWin10 = Environment.OSVersion.Version.Major == 10 && Environment.OSVersion.Version.Minor == 0 && Environment.OSVersion.Version.Build < 22000;
                int btnY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;

                if (!_buttonsExpanded)
                {
                    int totalButtonsWidth = 3 * BTN_SQUARE_SIZE + 2 * 10;
                    int startX = barRect.Right - totalButtonsWidth - 20;
                    Point expandPos = new Point(startX, btnY);
                    Point editPos = new Point(expandPos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point thirdPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);

                    DrawWin11Button(g, expandBtnRect, "展开", buttonIcons.ContainsKey("展开") ? buttonIcons["展开"] : null, ref expandBtnRect, expandPos, _expandPressed, "expand");
                    DrawWin11Button(g, editBtnRect, editMode ? "完成" : "编辑", buttonIcons.ContainsKey(editMode ? "完成" : "编辑") ? buttonIcons[editMode ? "完成" : "编辑"] : null, ref editBtnRect, editPos, _editPressed, "edit");
                    DrawWin11Button(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreen ? "缩小" : "全屏") ? buttonIcons[fullscreen ? "缩小" : "全屏"] : null, ref fullscreenBtnRect, thirdPos, _fullscreenPressed, "fullscreen");
                }
                else
                {
                    int totalButtonsWidth = 9 * BTN_SQUARE_SIZE + 8 * 10;
                    int startX = barRect.Right - totalButtonsWidth;
                    Point collapsePos = new Point(startX, btnY);
                    Point settingsPos = new Point(collapsePos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point rotatePos = new Point(settingsPos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point exportPos = new Point(rotatePos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point editPos = new Point(exportPos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point historyPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point fullscreenPos = new Point(historyPos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point minimizePos = new Point(fullscreenPos.X + BTN_SQUARE_SIZE + 10, btnY);
                    Point closePos = new Point(minimizePos.X + BTN_SQUARE_SIZE + 10, btnY);

                    DrawWin11Button(g, collapseBtnRect, "收起", buttonIcons.ContainsKey("收起") ? buttonIcons["收起"] : null, ref collapseBtnRect, collapsePos, _collapsePressed, "collapse");
                    DrawWin11Button(g, settingsBtnRect, "设置", buttonIcons.ContainsKey("设置") ? buttonIcons["设置"] : null, ref settingsBtnRect, settingsPos, _settingsPressed, "settings");
                    DrawWin11Button(g, rotateBtnRect, "轮换", buttonIcons.ContainsKey("轮换") ? buttonIcons["轮换"] : null, ref rotateBtnRect, rotatePos, _rotatePressed, "rotate");
                    DrawWin11Button(g, exportBtnRect, "导出", buttonIcons.ContainsKey("导出") ? buttonIcons["导出"] : null, ref exportBtnRect, exportPos, _exportPressed, "export");
                    DrawWin11Button(g, editBtnRect, editMode ? "完成" : "编辑", buttonIcons.ContainsKey(editMode ? "完成" : "编辑") ? buttonIcons[editMode ? "完成" : "编辑"] : null, ref editBtnRect, editPos, _editPressed, "edit");
                    DrawWin11Button(g, historyBtnRect, historyMode ? "返回" : "历史", buttonIcons.ContainsKey(historyMode ? "返回" : "历史") ? buttonIcons[historyMode ? "返回" : "历史"] : null, ref historyBtnRect, historyPos, _historyPressed, "history");
                    DrawWin11Button(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreen ? "缩小" : "全屏") ? buttonIcons[fullscreen ? "缩小" : "全屏"] : null, ref fullscreenBtnRect, fullscreenPos, _fullscreenPressed, "fullscreen");
                    DrawWin11Button(g, minimizeBtnRect, "最小化", buttonIcons.ContainsKey("最小化") ? buttonIcons["最小化"] : null, ref minimizeBtnRect, minimizePos, _minimizePressed, "minimize");
                    DrawWin11Button(g, closeBtnRect, "关闭", buttonIcons.ContainsKey("关闭") ? buttonIcons["关闭"] : null, ref closeBtnRect, closePos, _closePressed, "close");
                }
            }
        }

        private void DrawWin11Button(Graphics g, Rectangle rect, string text, Image icon, ref Rectangle targetRect, Point pos, bool pressed, string key)
        {
            targetRect = new Rectangle(pos, new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));
            bool isPressed = _buttonPressed.GetValueOrDefault(key) || pressed;

            float hoverProgress = _buttonHoverProgress.GetValueOrDefault(key, 0);

            int centerX = targetRect.X + targetRect.Width / 2;
            int centerY = targetRect.Y + targetRect.Height / 2;
            int maxRadius = targetRect.Width / 2;

            if (hoverProgress > 0)
            {
                float radiusGrowth = (float)Math.Pow(hoverProgress, 0.8);
                for (int i = 0; i < 4; i++)
                {
                    float radiusFactor = 0.6f + i * 0.15f;
                    float radius = maxRadius * radiusFactor * radiusGrowth;
                    int baseAlpha = 25 - i * 10;
                    if (baseAlpha < 0) baseAlpha = 0;
                    int alpha = (int)(baseAlpha * hoverProgress);
                    if (alpha > 0)
                    {
                        using (var brush = new SolidBrush(Color.FromArgb(alpha, 255, 255, 255)))
                        {
                            g.FillEllipse(brush, centerX - radius, centerY - radius, radius * 2, radius * 2);
                        }
                    }
                }
            }

            if (icon != null)
            {
                int targetHeight = (int)(targetRect.Height * 0.5);
                int targetWidth = (int)((float)icon.Width / icon.Height * targetHeight);
                int iconX = targetRect.Left + (targetRect.Width - targetWidth) / 2;
                int iconY = targetRect.Top + (targetRect.Height - targetHeight) / 2 - 7;
                g.DrawImage(icon, iconX, iconY, targetWidth, targetHeight);
            }

            using (var smallFont = new Font("微软雅黑", 8))
            {
                SizeF textSize = g.MeasureString(text, smallFont);
                float textX = targetRect.Left + (targetRect.Width - textSize.Width) / 2;
                float textY = targetRect.Bottom - textSize.Height - 1;
                g.DrawString(text, smallFont, Brushes.White, textX, textY);
            }
        }

        private void DrawDefaultArrow(Graphics g, Rectangle rect, bool left)
        {
            using (var brush = new SolidBrush(Color.DarkGray))
            {
                Point[] points;
                if (left)
                    points = new Point[] { new Point(rect.Right, rect.Top), new Point(rect.Right, rect.Bottom), new Point(rect.Left, rect.Top + rect.Height / 2) };
                else
                    points = new Point[] { new Point(rect.Left, rect.Top), new Point(rect.Left, rect.Bottom), new Point(rect.Right, rect.Top + rect.Height / 2) };
                g.FillPolygon(brush, points);
            }
        }

        private void DrawImageCentered(Graphics g, Image img, Rectangle container)
        {
            float scale = Math.Min((float)container.Width / img.Width, (float)container.Height / img.Height);
            int w = (int)(img.Width * scale);
            int h = (int)(img.Height * scale);
            int x = container.X + (container.Width - w) / 2;
            int y = container.Y + (container.Height - h) / 2;
            g.DrawImage(img, x, y, w, h);
        }

        // ------------------------------ 时间提醒与闪烁 ------------------------------
        private void CheckEveningClassStates()
        {
            if (!appConfig.ShowDueTime || appConfig.EveningClassTimes == null || appConfig.EveningClassTimes.Count == 0)
            {
                if (_activeEvenings.Count > 0 || _flashingEvenings.Count > 0 || _grayEvenings.Count > 0)
                {
                    _activeEvenings.Clear();
                    _flashingEvenings.Clear();
                    _grayEvenings.Clear();
                    StopFlashingIfNeeded();
                    Invalidate();
                }
                return;
            }

            DateTime now = DateTime.Now;
            var newActive = new List<string>();
            var newFlashing = new List<string>();
            var newGray = new List<string>();

            for (int i = 0; i < appConfig.EveningClassTimes.Count; i++)
            {
                var time = appConfig.EveningClassTimes[i];
                string eveningName = $"晚修{i + 1}";
                if (DateTime.TryParse(time.Start, out DateTime start) && DateTime.TryParse(time.End, out DateTime end))
                {
                    DateTime startToday = DateTime.Today.Add(start.TimeOfDay);
                    DateTime endToday = DateTime.Today.Add(end.TimeOfDay);
                    if (now >= startToday && now < endToday) newActive.Add(eveningName);
                    else if (now >= endToday.AddMinutes(-2) && now <= endToday.AddMinutes(2)) newFlashing.Add(eveningName);
                    else if (now > endToday.AddMinutes(2)) newGray.Add(eveningName);
                }
            }

            bool changed = !_activeEvenings.SequenceEqual(newActive) || !_flashingEvenings.SequenceEqual(newFlashing) || !_grayEvenings.SequenceEqual(newGray);
            if (changed)
            {
                _activeEvenings = newActive;
                _flashingEvenings = newFlashing;
                _grayEvenings = newGray;
                if (_flashingEvenings.Count > 0 && !_debugFlashing) StartFlashing();
                else if (_flashingEvenings.Count == 0 && !_debugFlashing) StopFlashingIfNeeded();
                Invalidate();
            }
        }

        private void StartFlashing() { if (!flashTimer.Enabled) { flashStartTime = DateTime.Now; flashTimer.Start(); } }
        private void StopFlashingIfNeeded() { if (_flashingEvenings.Count == 0 && !_debugFlashing) { flashTimer.Stop(); Invalidate(); } }
        private void FlashTimer_Tick(object sender, EventArgs e)
        {
            if (_debugFlashing && (DateTime.Now - _debugFlashStartTime).TotalSeconds > FLASH_DURATION) { _debugFlashing = false; StopFlashingIfNeeded(); }
            double angle = (DateTime.Now - flashStartTime).TotalMilliseconds / 500.0;
            flashStep = (int)((Math.Sin(angle) + 1) * 60);
            _laserOffset += 0.1f;
            if (_laserOffset >= 1f) _laserOffset -= 1f;
            Invalidate();
        }
        public void StartDebugFlashing() { _debugFlashing = true; _debugFlashStartTime = DateTime.Now; StartFlashing(); Invalidate(); }
        public void StopDebugFlashing() { _debugFlashing = false; StopFlashingIfNeeded(); }

        // ------------------------------ 提交时间下拉框管理 ------------------------------
        private void CreateTimeComboBoxes()
        {
            DestroyTimeComboBoxes();
            for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
            {
                string subject = currentSubjects[i];
                Rectangle virtualRect = GetDueTimeRect(i);
                Point screenLoc = MapToScreen(virtualRect.Location);
                Size screenSize = new Size((int)(virtualRect.Width * scaleFactor), (int)(virtualRect.Height * scaleFactor));
                string currentValue = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";

                var combo = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = screenLoc,
                    Size = new Size(screenSize.Width, 25),
                    Font = font22,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(40, 40, 40),
                    ForeColor = Color.White,
                    DrawMode = DrawMode.OwnerDrawFixed,
                    Tag = i
                };
                for (int j = 1; j <= appConfig.EveningClassCount; j++) combo.Items.Add($"晚修{j}");
                combo.Items.Add("无");
                int selectedIndex = combo.Items.IndexOf(currentValue);
                if (selectedIndex < 0) selectedIndex = appConfig.EveningClassCount;
                combo.SelectedIndex = selectedIndex;
                combo.DrawItem += TimeComboBox_DrawItem;
                combo.SelectedIndexChanged += TimeComboBox_SelectedIndexChanged;
                Controls.Add(combo);
                timeComboBoxes.Add(combo);
            }
        }

        private void DestroyTimeComboBoxes()
        {
            foreach (var combo in timeComboBoxes)
            {
                combo.SelectedIndexChanged -= TimeComboBox_SelectedIndexChanged;
                combo.DrawItem -= TimeComboBox_DrawItem;
                Controls.Remove(combo);
                combo.Dispose();
            }
            timeComboBoxes.Clear();
        }

        private void TimeComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var combo = sender as ComboBox;
            string text = combo.Items[e.Index].ToString();
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bgColor = selected ? Color.FromArgb(80, 80, 80) : Color.FromArgb(40, 40, 40);
            using (var bgBrush = new SolidBrush(bgColor))
                e.Graphics.FillRectangle(bgBrush, e.Bounds);
            using (var textBrush = new SolidBrush(Color.White))
                e.Graphics.DrawString(text, combo.Font, textBrush, e.Bounds.Left + 5, e.Bounds.Top + 2);
            e.DrawFocusRectangle();
        }

        private void TimeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var combo = sender as ComboBox;
            int subjectIndex = (int)combo.Tag;
            string subject = currentSubjects[subjectIndex];
            homeworkData.DueTimes[subject] = combo.SelectedItem?.ToString() ?? "";
            SaveHomeworkData();
        }

        // ------------------------------ 动态控件管理 ------------------------------
        private void DestroyAllDynamicControls()
        {
            if (inlineEditControl != null && !inlineEditControl.IsDisposed)
            {
                try
                {
                    Controls.Remove(inlineEditControl);
                    inlineEditControl.Dispose();
                }
                catch { }
                inlineEditControl = null;
                editingSubjectIndex = -1;
                currentEditType = EditFieldType.None;
            }

            if (timeComboBoxes != null)
            {
                foreach (var combo in timeComboBoxes.ToList())
                {
                    if (combo != null && !combo.IsDisposed)
                    {
                        try
                        {
                            combo.SelectedIndexChanged -= TimeComboBox_SelectedIndexChanged;
                            combo.DrawItem -= TimeComboBox_DrawItem;
                            Controls.Remove(combo);
                            combo.Dispose();
                        }
                        catch { }
                    }
                }
                timeComboBoxes.Clear();
            }
        }

        private void RecreateDynamicControls()
        {
            if (EditMode && !rotationMode && !IsDisposed)
            {
                CreateTimeComboBoxes();
                // 编辑模式下纯文本编辑框已在 EditMode setter 中创建，此处不需要重复创建
            }
            else if (!EditMode && !rotationMode)
            {
                if (_useWebView2)
                {
                    if (cardWebViews.Count == 0)
                        RecreateWebView2ForGrid();
                    else
                        UpdateWebView2Positions();
                }
                // GDI+ 模式无需重建，重绘即可
            }
            else if (rotationMode && _useWebView2)
            {
                if (rotationWebView == null)
                    CreateRotationWebView();
                else
                    UpdateWebView2Positions();
            }
        }

        // ------------------------------ 卡片飞入动画 ------------------------------
        private void StartCardFlyInAnimation()
        {
            _isFlyingIn = true;
            _suspendMouseGlow = true;
            _mouseGlowAlpha = 0;
            Invalidate();

            _cardAnimations.Clear();
            int cols = 3;
            double totalDelay = 0.5;
            double colDelay = 0.12;
            double duration = 0.45;

            _flyInAnimationsRemaining = gridRects.Count;

            for (int i = 0; i < gridRects.Count; i++)
            {
                int col = i % cols;
                int row = i / cols;
                double delay = totalDelay + col * colDelay + row * 0.02;

                var anim = new Animation
                {
                    Duration = TimeSpan.FromSeconds(duration),
                    StartDelay = TimeSpan.FromSeconds(delay),
                    Tag = i
                };
                int index = i;
                anim.OnUpdate = (p) =>
                {
                    if (index < gridRects.Count)
                        Invalidate(gridRects[index]);
                };
                anim.OnComplete = () =>
                {
                    _flyInAnimationsRemaining--;
                    if (_flyInAnimationsRemaining == 0)
                    {
                        _isFlyingIn = false;
                        _suspendMouseGlow = false;
                        if (!_isWin10)
                        {
                            var fadeIn = new Animation
                            {
                                Duration = TimeSpan.FromMilliseconds(500),
                                Tag = "mouseGlowFadeIn"
                            };
                            fadeIn.OnUpdate = (p) =>
                            {
                                _mouseGlowAlpha = (float)p;
                                Invalidate();
                            };
                            fadeIn.OnComplete = () =>
                            {
                                _mouseGlowAlpha = 1;
                                _mouseGlowFadeAnimation = null;
                                Invalidate();
                            };
                            fadeIn.Start();
                            _animations.Add(fadeIn);
                            _mouseGlowFadeAnimation = fadeIn;
                        }
                        else
                        {
                            _mouseGlowAlpha = 0;
                        }
                    }
                };
                anim.Start();
                _cardAnimations[index] = anim;
                _animations.Add(anim);
            }
        }

        // ------------------------------ 窗体生命周期 ------------------------------
        private void OnLoad(object sender, EventArgs e)
        {
            ApplyBackgroundEffect(appConfig.BackgroundEffect);
            if (appConfig.UpdatePending == 1)
            {
                string upgradePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "HomeworkViewerUpgrader");
                if (Directory.Exists(upgradePath)) try { Directory.Delete(upgradePath, true); } catch { }
                appConfig.UpdatePending = 0;
                appConfig.Save();
            }
            var version = Environment.OSVersion.Version;
            if (version.Major == 10 && version.Minor == 0 && version.Build < 22000)
            {
                FormBorderStyle = FormBorderStyle.None;
                WindowState = FormWindowState.Maximized;
                fullscreen = true;
                UpdateScale();
            }

            InitializeDrawerPanel();
            StartCardFlyInAnimation();

            // 初始化时，若为 WebView2 模式且未初始化，等待初始化完成后刷新内容
            if (_useWebView2)
            {
                _ = Task.Run(async () =>
                {
                    while (!webView2Initialized)
                        await Task.Delay(100);
                    this.Invoke((MethodInvoker)(() =>
                    {
                        if (!rotationMode && !EditMode)
                            RefreshAllWebView2Content();
                    }));
                });
            }
            else
            {
                // GDI+ 模式：直接触发重绘
                Invalidate();
            }
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e) { }
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            rotationTimer?.Dispose();
            timeCheckTimer?.Dispose();
            flashTimer?.Dispose();
            scrollTimer?.Dispose();
            _animationTimer?.Dispose();
            font12?.Dispose(); font20?.Dispose(); font24?.Dispose(); font30?.Dispose(); font22?.Dispose(); font36?.Dispose();
            hintFont?.Dispose(); buttonFont?.Dispose(); fontSmall?.Dispose();
            RED_SEMI?.Dispose(); ORANGE_SEMI?.Dispose(); GREEN_SEMI?.Dispose(); BLUE_SEMI?.Dispose(); PURPLE_SEMI?.Dispose(); DARKORANGE_SEMI?.Dispose();
            DestroyAllWebView2();
            DestroyPlainTextEditors();
            base.OnFormClosed(e);
        }

        private void ShowExportDialog()
        {
            // 导出对话框实现（可保留原有逻辑）
        }
    }

    // ------------------------------ 动画类 ------------------------------
    public class Animation
    {
        public double Progress { get; set; } = 0;
        public TimeSpan Duration { get; set; } = TimeSpan.FromMilliseconds(300);
        public TimeSpan StartDelay { get; set; } = TimeSpan.Zero;
        public DateTime StartTime { get; private set; }
        public bool IsRunning { get; private set; }
        public bool IsLooping { get; set; } = false;
        public Action<double> OnUpdate { get; set; }
        public Action OnComplete { get; set; }
        public object Tag { get; set; }

        public void Start()
        {
            StartTime = DateTime.Now + StartDelay;
            IsRunning = true;
            Progress = 0;
        }

        public void Update()
        {
            if (!IsRunning) return;
            double elapsed = (DateTime.Now - StartTime).TotalMilliseconds;
            if (elapsed < 0) return;
            double newProgress = Math.Min(1.0, elapsed / Duration.TotalMilliseconds);
            Progress = newProgress;
            OnUpdate?.Invoke(Progress);
            if (newProgress >= 1.0)
            {
                if (IsLooping)
                {
                    StartTime = DateTime.Now + StartDelay;
                    Progress = 0;
                }
                else
                {
                    IsRunning = false;
                    OnComplete?.Invoke();
                }
            }
        }
    }
}