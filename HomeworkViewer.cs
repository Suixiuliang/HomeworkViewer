#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace HomeworkViewer
{
    public class HomeworkViewer : Form
    {
        // 常量
        private readonly Size VIRTUAL_SIZE = new Size(1200, 675);
        private const int BTN_SQUARE_SIZE = 46;
        private readonly Point FULLSCREEN_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10, 13);
        private readonly Point HISTORY_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10, 13);
        private readonly Point EDIT_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10, 13);
        private readonly Point ROTATE_BTN_POS = new Point(1200 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10, 13);
        private Point SETTINGS_BTN_POS;
        private readonly Point UNSUBMITTED_BTN_POS;
        private readonly Rectangle GRID_AREA = new Rectangle(0, 47, 1200, 631);
        private const int GRID_COLS = 3, GRID_ROWS = 2, GRID_PADDING = 20, GRID_BORDER = 1;
        private const int ROUND_RADIUS = 5;

        // 顶部黄色 Bar 参数
        private const int BAR_HEIGHT = 46;
        private const int BAR_Y = 0;
        private const int BAR_PADDING = 20;

        // Mica 效果相关
        private bool _micaEnabled = false;
        private int _backgroundAlpha = 32;

        // 状态
        private bool fullscreen = false;
        private bool historyMode = false;
        private bool editMode = false;
        private bool rotationMode = false;
        private bool unsubmittedMode = false;
        private int rotationIndex = 0;
        private int unsubmittedPage = 0;
        private Timer rotationTimer;
        private const int rotationInterval = 10000;

        // 内联编辑
        private Control inlineEditControl;
        private int editingSubjectIndex = -1;
        private enum EditFieldType { None, Subject, DueTime, Unsubmitted }
        private EditFieldType currentEditType = EditFieldType.None;

        // 数据
        private HomeworkData homeworkData = new HomeworkData();
        private DateTime currentDate = DateTime.Now;
        private DateTime? historyDate = null;

        // 按钮矩形
        private Rectangle rotateBtnRect, editBtnRect, historyBtnRect, fullscreenBtnRect, backBtnRect, leftArrowRect, rightArrowRect, settingsBtnRect, unsubmittedBtnRect, minimizeBtnRect, closeBtnRect, exportBtnRect;
        private Rectangle expandBtnRect, collapseBtnRect;

        // 网格矩形
        private List<Rectangle> gridRects = new List<Rectangle>();
        private Rectangle[,] fieldRects;

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

        // 缩放相关
        private float scaleFactor = 1.0f;
        private Point offset = Point.Empty;
        private bool _resizing = false;

        // 图片资源
        private Image buttonImage, historyBtnImage, backBtnImage, leftArrowImage, rightArrowImage;
        private Dictionary<string, Image> buttonIcons = new Dictionary<string, Image>();

        // 配置
        private AppConfig appConfig;

        // 字体缩放因子
        private float[] fontScales = { 0.8f, 1.0f, 1.2f };
        private string[] currentSubjects;

        // 保存最小化前的窗口大小和状态
        private Size _savedClientSize;
        private FormWindowState _savedWindowState;

        // 按钮按下状态
        private bool _settingsPressed = false;
        private bool _rotatePressed = false;
        private bool _editPressed = false;
        private bool _historyPressed = false;
        private bool _fullscreenPressed = false;
        private bool _backPressed = false;
        private bool _unsubmittedPressed = false;
        private bool _minimizePressed = false;
        private bool _closePressed = false;
        private bool _expandPressed = false;
        private bool _collapsePressed = false;
        private bool _exportPressed = false;

        // 收纳状态
        private bool _buttonsExpanded = false;

        // 模式切换下拉框
        private ComboBox modeComboBox;

        // 编辑模式下的提交时间下拉框列表
        private List<ComboBox> timeComboBoxes = new List<ComboBox>();

        // 时间提醒相关
        private Timer timeCheckTimer;
        private List<string> _activeEvenings = new List<string>();
        private List<string> _flashingEvenings = new List<string>();
        private List<string> _grayEvenings = new List<string>();
        private bool _previousFlashingState = false;

        // 闪烁相关
        private Timer flashTimer;
        private int flashStep = 0;
        private bool _debugFlashing = false;
        private DateTime _debugFlashStartTime;
        private DateTime flashStartTime;
        private const int FLASH_DURATION = 300;
        private const int FLASH_INTERVAL = 100;
        private float _laserOffset = 0f;

        // 滚动相关
        private Timer scrollTimer;
        private Dictionary<int, float> scrollOffsets = new Dictionary<int, float>();
        private Dictionary<int, bool> scrollPaused = new Dictionary<int, bool>();
        private Dictionary<int, DateTime> pauseStartTime = new Dictionary<int, DateTime>();
        private const int SCROLL_PAUSE_SECONDS = 3;

        // 未交名单每页显示行数
        private const int UNSUBMITTED_ROWS_PER_PAGE = 3;

        // 背景效果相关
        private string _currentBackgroundEffect = "Mica";
        private bool _isWin10OrAbove = false;
        private bool _sizing = false;

        // Markdown 模式滚动
        private float _markdownScrollOffset = 0f;
        private bool _markdownScrollPaused = false;
        private DateTime _markdownPauseStart = DateTime.Now;
        private float _markdownTotalHeight = 0;

        // 行列调整相关
        private ResizeHelper _resizeHelper;
        private int[] _rowHeights = new int[2];   // 最多2行
        private int[] _colWidths = new int[3];    // 固定3列
        private bool _isResizing = false;
        private int _resizeTargetRow = -1;
        private int _resizeTargetCol = -1;
        private int _resizeStartX, _resizeStartY;
        private int _originalHeight, _originalWidth;

        // 导出相关
        private Button exportBtn;

        public HomeworkViewer()
        {
            SETTINGS_BTN_POS = new Point(ROTATE_BTN_POS.X - BTN_SQUARE_SIZE - 10, 13);
            UNSUBMITTED_BTN_POS = new Point(ROTATE_BTN_POS.X - BTN_SQUARE_SIZE - 10 - BTN_SQUARE_SIZE - 10, 13);

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
            _backgroundAlpha = (int)(appConfig.BackgroundOpacity / 100f * 255);
            ApplyFontSettings();
            UpdateSubjectsByMode();
            CalculateGridRects();
            LoadHomeworkData(currentDate);
            LoadImages();

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

            this.MouseClick += OnMouseClick;
            this.MouseDown += OnMouseDown;
            this.MouseUp += OnMouseUp;
            this.MouseLeave += OnMouseLeave;
            this.Resize += OnResize;
            this.Load += OnLoad;
            this.Activated += OnActivated;
            this.FormClosing += OnFormClosing;
            this.KeyDown += OnKeyDown;

            InitializeModeComboBox();

            CheckWindowsVersion();
            ApplyBackgroundEffect(appConfig.BackgroundEffect);

            CheckEveningClassStates();

            // 初始化行列调整
            _rowHeights[0] = _rowHeights[1] = 0; // 0表示自动
            _colWidths[0] = _colWidths[1] = _colWidths[2] = 0;
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
            {
                ToggleFullscreen();
                e.Handled = true;
            }
        }

        private void CheckWindowsVersion()
        {
            OperatingSystem os = Environment.OSVersion;
            Version v = os.Version;
            if (v.Major >= 10)
            {
                _isWin10OrAbove = true;
            }
        }

        public void ApplyBackgroundEffect(string effect)
        {
            _currentBackgroundEffect = effect;
            try
            {
                switch (effect)
                {
                    case "Mica":
                        EnableMica();
                        break;
                    case "Acrylic":
                        EnableAcrylic();
                        break;
                    case "Aero":
                        EnableAero();
                        break;
                    default:
                        EnableMica();
                        break;
                }
            }
            catch { }
        }

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
                    else
                    {
                        EnableAcrylicFallback();
                    }
                }
                else
                {
                    EnableAcrylicFallback();
                }
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

        private void EnableAcrylicFallback()
        {
            EnableAcrylic();
        }

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
                }
                else if (m.Msg == WM_EXITSIZEMOVE)
                {
                    _sizing = false;
                    UpdateScale();
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
                {
                    this.ClientSize = _savedClientSize;
                }
                UpdateScale();
                if (!_sizing)
                {
                    Invalidate();
                }
                if (!_micaEnabled)
                {
                    ApplyBackgroundEffect(_currentBackgroundEffect);
                }
            }

            base.WndProc(ref m);
        }

        private void OnResize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                _savedClientSize = this.ClientSize;
                _savedWindowState = this.WindowState;
                return;
            }

            if (_resizing || fullscreen || WindowState == FormWindowState.Maximized) return;

            _resizing = true;
            float targetRatio = (float)VIRTUAL_SIZE.Width / VIRTUAL_SIZE.Height;
            int newWidth = this.ClientSize.Width;
            int newHeight = (int)(newWidth / targetRatio);
            this.ClientSize = new Size(newWidth, newHeight);
            _resizing = false;

            UpdateScale();
            if (EditMode && !unsubmittedMode)
            {
                CreateTimeComboBoxes();
            }
            if (!_isWin10OrAbove || !_sizing)
            {
                Invalidate();
            }
        }

        public void ApplySettings(AppConfig newConfig)
        {
            appConfig = newConfig;
            ApplyFontSettings();
            UpdateSubjectsByMode();
            _backgroundAlpha = (int)(appConfig.BackgroundOpacity / 100f * 255);
            CheckEveningClassStates();
            if (EditMode)
            {
                CreateTimeComboBoxes();
            }
            ApplyBackgroundEffect(appConfig.BackgroundEffect);
            Invalidate();
        }

        private void ApplyFontSettings()
        {
            font12?.Dispose();
            font20?.Dispose();
            font24?.Dispose();
            font30?.Dispose();
            font22?.Dispose();
            font36?.Dispose();
            hintFont?.Dispose();
            buttonFont?.Dispose();
            fontSmall?.Dispose();

            float scale = fontScales[appConfig.FontSizeLevel];
            string fontName = appConfig.FontFamily;
            if (appConfig.IsCustomFont)
            {
                // 使用自定义字体
                Font customFont = FontManager.GetCustomFont(20 * scale);
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
            if (subjectModes.ContainsKey(appConfig.LastMode))
                currentSubjects = subjectModes[appConfig.LastMode];
            else
                currentSubjects = subjectModes["大理"];
            CalculateGridRects();
        }

        // ---------- Mica API ----------
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

        // ---------- 网格计算 ----------
        private void CalculateGridRects()
        {
            gridRects.Clear();
            int subjectCount = currentSubjects.Length;
            int cols = 3;
            int rows = (subjectCount + cols - 1) / cols;

            float areaWidth = GRID_AREA.Width;
            float areaHeight = GRID_AREA.Height;

            // 如果用户自定义了行高/列宽，使用自定义值
            float rectWidth, rectHeight;
            if (appConfig.ColumnWidth > 0)
                rectWidth = appConfig.ColumnWidth;
            else
                rectWidth = (areaWidth - (cols + 1) * GRID_PADDING) / cols;

            if (appConfig.RowHeight > 0)
                rectHeight = appConfig.RowHeight;
            else
                rectHeight = (areaHeight - (rows + 1) * GRID_PADDING) / rows;

            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < cols; col++)
                {
                    int index = row * cols + col;
                    if (index >= subjectCount) break;
                    int x = GRID_AREA.Left + GRID_PADDING + col * (int)(rectWidth + GRID_PADDING);
                    int y = GRID_AREA.Top + GRID_PADDING + row * (int)(rectHeight + GRID_PADDING);
                    gridRects.Add(new Rectangle(x, y, (int)rectWidth, (int)rectHeight));
                }
            }
            fieldRects = new Rectangle[gridRects.Count, 1];
        }

        private void LoadHomeworkData(DateTime date)
        {
            homeworkData = HomeworkData.Load(date);

            if (homeworkData.DueTimes.Count == 0)
            {
                foreach (string subject in currentSubjects)
                {
                    string defaultDue = appConfig.EveningClassCount >= 3 ? "晚修3" : "无";
                    homeworkData.DueTimes[subject] = defaultDue;
                }
                SaveHomeworkData();
            }

            if (EditMode && !unsubmittedMode)
            {
                CreateTimeComboBoxes();
            }
        }

        private void SaveHomeworkData()
        {
            DateTime saveDate = historyMode && historyDate.HasValue ? historyDate.Value : currentDate;
            homeworkData.Save(saveDate);
        }

        // ---------- 全屏切换 ----------
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
            Invalidate();
        }

        // ---------- 缩放相关 ----------
        private void UpdateScale()
        {
            Size clientSize = this.ClientSize;
            if (clientSize.Width == 0 || clientSize.Height == 0) return;

            float scaleW = (float)clientSize.Width / VIRTUAL_SIZE.Width;
            float scaleH = (float)clientSize.Height / VIRTUAL_SIZE.Height;
            scaleFactor = Math.Min(scaleW, scaleH);
            Size scaledVirtual = new Size((int)(VIRTUAL_SIZE.Width * scaleFactor), (int)(VIRTUAL_SIZE.Height * scaleFactor));
            offset = new Point((clientSize.Width - scaledVirtual.Width) / 2, (clientSize.Height - scaledVirtual.Height) / 2);
        }

        private Point MapToVirtual(Point screenPt)
        {
            if (scaleFactor == 0) return screenPt;
            int vx = (int)((screenPt.X - offset.X) / scaleFactor);
            int vy = (int)((screenPt.Y - offset.Y) / scaleFactor);
            return new Point(vx, vy);
        }

        private Point MapToScreen(Point virtualPt)
        {
            if (scaleFactor == 0) return virtualPt;
            int sx = (int)(virtualPt.X * scaleFactor + offset.X);
            int sy = (int)(virtualPt.Y * scaleFactor + offset.Y);
            return new Point(sx, sy);
        }

        // ---------- 鼠标事件 ----------
        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            Point virtualPos = MapToVirtual(e.Location);
            int x = virtualPos.X, y = virtualPos.Y;

            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _unsubmittedPressed = _minimizePressed = _closePressed = _expandPressed = _collapsePressed = _exportPressed = false;

            if (rotationMode || unsubmittedMode)
            {
                if (backBtnRect.Contains(x, y))
                    _backPressed = true;
                if (editBtnRect.Contains(x, y))
                    _editPressed = true;
                if (historyBtnRect.Contains(x, y))
                    _historyPressed = true;
                if (fullscreenBtnRect.Contains(x, y))
                    _fullscreenPressed = true;
                if (exportBtnRect.Contains(x, y))
                    _exportPressed = true;
            }
            else
            {
                if (!_buttonsExpanded)
                {
                    if (expandBtnRect.Contains(x, y))
                        _expandPressed = true;
                    else if (editBtnRect.Contains(x, y))
                        _editPressed = true;
                    else
                    {
                        bool isWin10 = Environment.OSVersion.Version.Major == 10 && Environment.OSVersion.Version.Minor == 0 && Environment.OSVersion.Version.Build < 22000;
                        if (isWin10)
                        {
                            if (minimizeBtnRect.Contains(x, y))
                                _minimizePressed = true;
                        }
                        else
                        {
                            if (fullscreenBtnRect.Contains(x, y))
                                _fullscreenPressed = true;
                        }
                    }
                }
                else
                {
                    if (collapseBtnRect.Contains(x, y))
                        _collapsePressed = true;
                    else if (settingsBtnRect.Contains(x, y))
                        _settingsPressed = true;
                    else if (rotateBtnRect.Contains(x, y))
                        _rotatePressed = true;
                    else if (unsubmittedBtnRect.Contains(x, y))
                        _unsubmittedPressed = true;
                    else if (exportBtnRect.Contains(x, y))
                        _exportPressed = true;
                    else if (editBtnRect.Contains(x, y))
                        _editPressed = true;
                    else if (historyBtnRect.Contains(x, y))
                        _historyPressed = true;
                    else if (fullscreenBtnRect.Contains(x, y))
                        _fullscreenPressed = true;
                    else if (minimizeBtnRect.Contains(x, y))
                        _minimizePressed = true;
                    else if (closeBtnRect.Contains(x, y))
                        _closePressed = true;
                }
            }

            // 检测是否在网格线附近进行行列调整
            if (!rotationMode && !unsubmittedMode && !editMode)
            {
                for (int i = 0; i < gridRects.Count; i++)
                {
                    var rect = gridRects[i];
                    int col = i % 3;
                    int row = i / 3;
                    // 检测右边界
                    if (Math.Abs(x - rect.Right) < 5 && col < 2)
                    {
                        _isResizing = true;
                        _resizeTargetCol = col;
                        _resizeStartX = x;
                        _originalWidth = rect.Width;
                        break;
                    }
                    // 检测下边界
                    if (Math.Abs(y - rect.Bottom) < 5 && row < 1)
                    {
                        _isResizing = true;
                        _resizeTargetRow = row;
                        _resizeStartY = y;
                        _originalHeight = rect.Height;
                        break;
                    }
                }
            }

            Invalidate();
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            if (_isResizing)
            {
                _isResizing = false;
                // 保存调整后的尺寸到配置
                if (_resizeTargetCol >= 0)
                {
                    int newWidth = _originalWidth + (e.X - _resizeStartX);
                    if (newWidth > 100)
                        appConfig.ColumnWidth = newWidth;
                }
                if (_resizeTargetRow >= 0)
                {
                    int newHeight = _originalHeight + (e.Y - _resizeStartY);
                    if (newHeight > 80)
                        appConfig.RowHeight = newHeight;
                }
                appConfig.Save();
                CalculateGridRects();
                Invalidate();
                _resizeTargetCol = -1;
                _resizeTargetRow = -1;
            }

            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _unsubmittedPressed = _minimizePressed = _closePressed = _expandPressed = _collapsePressed = _exportPressed = false;
            Invalidate();
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _unsubmittedPressed = _minimizePressed = _closePressed = _expandPressed = _collapsePressed = _exportPressed = false;
            Invalidate();
        }

        private void OnActivated(object sender, EventArgs e)
        {
            UpdateScale();
            Invalidate();
        }

        // ---------- 辅助方法 ----------
        private GraphicsPath CreateRoundedRectPath(Rectangle rect, int radius, bool bottomOnly = false)
        {
            GraphicsPath path = new GraphicsPath();
            if (radius <= 0)
            {
                path.AddRectangle(rect);
                return path;
            }

            if (bottomOnly)
            {
                int x = rect.X;
                int y = rect.Y;
                int w = rect.Width;
                int h = rect.Height;
                int r = radius;
                int d = r * 2;

                path.AddLine(x, y, x + w, y);
                path.AddLine(x + w, y, x + w, y + h - r);
                path.AddArc(new Rectangle(x + w - d, y + h - d, d, d), 0, 90);
                path.AddLine(x + w - r, y + h, x + r, y + h);
                path.AddArc(new Rectangle(x, y + h - d, d, d), 90, 90);
                path.AddLine(x, y + h - r, x, y);
            }
            else
            {
                int diameter = radius * 2;
                Rectangle arcRect = new Rectangle(rect.Location, new Size(diameter, diameter));

                path.AddArc(arcRect, 180, 90);
                arcRect.X = rect.Right - diameter;
                path.AddArc(arcRect, 270, 90);
                arcRect.Y = rect.Bottom - diameter;
                path.AddArc(arcRect, 0, 90);
                arcRect.X = rect.Left;
                path.AddArc(arcRect, 90, 90);
                path.CloseFigure();
            }
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
            using (Graphics g = this.CreateGraphics())
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
            if (height < 25)
            {
                minY = maxY - 25;
                height = 25;
            }
            return new Rectangle((int)startX, minY, (int)totalWidth + 5, height);
        }

        // ---------- 模式切换下拉框 ----------
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
            this.Controls.Add(modeComboBox);
        }

        private void ModeComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            ComboBox combo = sender as ComboBox;
            string text = combo.Items[e.Index].ToString();

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bgColor = isSelected ? Color.FromArgb(80, 80, 80) : Color.FromArgb(40, 40, 40);
            using (SolidBrush bgBrush = new SolidBrush(bgColor))
            {
                e.Graphics.FillRectangle(bgBrush, e.Bounds);
            }

            using (SolidBrush textBrush = new SolidBrush(Color.White))
            {
                e.Graphics.DrawString(text, combo.Font, textBrush, e.Bounds.Left + 5, e.Bounds.Top + 2);
            }

            e.DrawFocusRectangle();
        }

        private void ModeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string newMode = modeComboBox.SelectedItem.ToString();
            if (newMode == appConfig.LastMode) return;

            appConfig.LastMode = newMode;
            appConfig.Save();
            UpdateSubjectsByMode();
            if (EditMode && !unsubmittedMode)
            {
                CreateTimeComboBoxes();
            }
            Invalidate();
        }

        // ---------- 鼠标点击 ----------
        private void OnMouseClick(object sender, MouseEventArgs e)
        {
            Point virtualPos = MapToVirtual(e.Location);
            int x = virtualPos.X, y = virtualPos.Y;

            if (editingSubjectIndex != -1)
            {
                FinishInlineEdit();
            }

            if (rotationMode)
            {
                if (backBtnRect.Contains(x, y))
                {
                    rotationMode = false;
                    rotationTimer.Stop();
                    Invalidate();
                }
                else if (leftArrowRect.Contains(x, y))
                {
                    RotateManual(-1);
                }
                else if (rightArrowRect.Contains(x, y))
                {
                    RotateManual(1);
                }
                else if (fullscreenBtnRect.Contains(x, y))
                {
                    if (EditMode) EditMode = false;
                    ToggleFullscreen();
                }
                else if (exportBtnRect.Contains(x, y))
                {
                    ShowExportDialog();
                }
            }
            else if (unsubmittedMode)
            {
                if (backBtnRect.Contains(x, y))
                {
                    unsubmittedMode = false;
                    Invalidate();
                }
                else if (leftArrowRect.Contains(x, y))
                {
                    int totalPages = (currentSubjects.Length + UNSUBMITTED_ROWS_PER_PAGE - 1) / UNSUBMITTED_ROWS_PER_PAGE;
                    unsubmittedPage = (unsubmittedPage - 1 + totalPages) % totalPages;
                    Invalidate();
                }
                else if (rightArrowRect.Contains(x, y))
                {
                    int totalPages = (currentSubjects.Length + UNSUBMITTED_ROWS_PER_PAGE - 1) / UNSUBMITTED_ROWS_PER_PAGE;
                    unsubmittedPage = (unsubmittedPage + 1) % totalPages;
                    Invalidate();
                }
                else if (editBtnRect.Contains(x, y))
                {
                    if (EditMode)
                    {
                        SaveHomeworkData();
                    }
                    EditMode = !EditMode;
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
                }
                else if (fullscreenBtnRect.Contains(x, y))
                {
                    if (EditMode) EditMode = false;
                    ToggleFullscreen();
                }
                else if (exportBtnRect.Contains(x, y))
                {
                    ShowExportDialog();
                }
                else if (EditMode)
                {
                    Rectangle bigRect = new Rectangle(150, 100, 900, 475);
                    int startIndex = unsubmittedPage * UNSUBMITTED_ROWS_PER_PAGE;
                    int endIndex = Math.Min(startIndex + UNSUBMITTED_ROWS_PER_PAGE, currentSubjects.Length);
                    int rowHeight = (bigRect.Height - 120) / UNSUBMITTED_ROWS_PER_PAGE;
                    int baseY = bigRect.Top + 100;

                    for (int i = startIndex; i < endIndex; i++)
                    {
                        int rowY = baseY + (i - startIndex) * rowHeight;
                        Rectangle contentRect = new Rectangle(bigRect.Left + 200, rowY, bigRect.Width - 250, rowHeight);

                        if (contentRect.Contains(x, y))
                        {
                            StartInlineEdit(i, EditFieldType.Unsubmitted, contentRect);
                            return;
                        }
                    }
                }
            }
            else
            {
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
                        if (EditMode)
                        {
                            SaveHomeworkData();
                        }
                        EditMode = !EditMode;
                        return;
                    }
                    else
                    {
                        if (isWin10)
                        {
                            if (minimizeBtnRect.Contains(x, y))
                            {
                                this.WindowState = FormWindowState.Minimized;
                                return;
                            }
                        }
                        else
                        {
                            if (fullscreenBtnRect.Contains(x, y))
                            {
                                if (EditMode) EditMode = false;
                                ToggleFullscreen();
                                return;
                            }
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
                        bool hasContent = false;
                        foreach (var subj in currentSubjects)
                        {
                            if (homeworkData.Subjects.ContainsKey(subj) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[subj]))
                            { hasContent = true; break; }
                        }
                        if (!hasContent)
                        {
                            MessageBox.Show("所有科目都没有作业内容，无法进入轮播模式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        rotationMode = true;
                        rotationIndex = 0;
                        rotationTimer.Start();
                        Invalidate();
                        return;
                    }
                    else if (unsubmittedBtnRect.Contains(x, y))
                    {
                        unsubmittedMode = true;
                        unsubmittedPage = 0;
                        if (EditMode)
                        {
                            EditMode = false;
                        }
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
                        if (EditMode)
                        {
                            SaveHomeworkData();
                        }
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
                        this.WindowState = FormWindowState.Minimized;
                        return;
                    }
                    else if (closeBtnRect.Contains(x, y))
                    {
                        Application.Exit();
                        return;
                    }
                    else if (!isWin10 && fullscreenBtnRect.Contains(x, y))
                    {
                        if (EditMode) EditMode = false;
                        ToggleFullscreen();
                        return;
                    }
                }

                if (EditMode && !rotationMode && !unsubmittedMode)
                {
                    for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
                    {
                        string subject = currentSubjects[i];

                        Rectangle subjectArea = new Rectangle(
                            gridRects[i].Left + 10,
                            gridRects[i].Top + 60,
                            gridRects[i].Width - 20,
                            gridRects[i].Height - 70);
                        if (subjectArea.Contains(x, y))
                        {
                            StartInlineEdit(i, EditFieldType.Subject, subjectArea);
                            return;
                        }
                    }
                }
            }
        }

        private void ShowExportDialog()
        {
            using (var dlg = new Form())
            {
                dlg.Text = "导出作业";
                dlg.Size = new Size(400, 350);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.BackColor = Color.FromArgb(45, 45, 48);
                dlg.ForeColor = Color.White;

                CheckBox chkIncludeUnsubmitted = new CheckBox { Text = "包含未交名单", Location = new Point(20, 20), AutoSize = true, ForeColor = Color.White };
                CheckBox chkRangeMode = new CheckBox { Text = "日期范围", Location = new Point(20, 50), AutoSize = true, ForeColor = Color.White };
                DateTimePicker dtpStart = new DateTimePicker { Format = DateTimePickerFormat.Short, Location = new Point(150, 48), Width = 120, Visible = false };
                Label lblTo = new Label { Text = "至", Location = new Point(280, 52), AutoSize = true, ForeColor = Color.White, Visible = false };
                DateTimePicker dtpEnd = new DateTimePicker { Format = DateTimePickerFormat.Short, Location = new Point(310, 48), Width = 120, Visible = false };
                ComboBox cmbFormat = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = new Point(20, 100),
                    Width = 150,
                    BackColor = Color.FromArgb(64, 64, 64),
                    ForeColor = Color.White
                };
                cmbFormat.Items.AddRange(new object[] { "txt", "pdf", "jpg", "html" });
                cmbFormat.SelectedIndex = 0;

                Button btnExport = new Button { Text = "导出", Location = new Point(20, 150), Width = 80, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
                Button btnCancel = new Button { Text = "取消", Location = new Point(110, 150), Width = 80, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };
                Button btnShare = new Button { Text = "分享", Location = new Point(200, 150), Width = 80, BackColor = Color.FromArgb(64, 64, 64), ForeColor = Color.White };

                chkRangeMode.CheckedChanged += (s, e) =>
                {
                    bool range = chkRangeMode.Checked;
                    dtpStart.Visible = range;
                    lblTo.Visible = range;
                    dtpEnd.Visible = range;
                };

                btnExport.Click += async (s, e) =>
                {
                    string format = cmbFormat.SelectedItem.ToString();
                    string fileName = $"作业_{DateTime.Now:yyyyMMdd_HHmmss}.{format}";
                    string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);

                    try
                    {
                        if (chkRangeMode.Checked)
                        {
                            // 范围导出
                            var start = dtpStart.Value.Date;
                            var end = dtpEnd.Value.Date;
                            var dates = homeworkData.GetDatesInRange(start, end);
                            if (dates.Count == 0)
                            {
                                MessageBox.Show("指定范围内没有作业数据。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                            // 简单处理：导出所有日期的 txt 合并
                            var sb = new System.Text.StringBuilder();
                            foreach (var date in dates)
                            {
                                var data = HomeworkData.Load(date);
                                sb.AppendLine($"===== {date:yyyy年MM月dd日} =====");
                                sb.AppendLine();
                                // 这里可以复用 ExportHelper 的单日逻辑
                            }
                            File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8);
                        }
                        else
                        {
                            // 导出当前日期
                            var data = homeworkData;
                            if (format == "txt")
                                ExportHelper.ExportToTxt(data, currentDate, filePath, chkIncludeUnsubmitted.Checked);
                            else if (format == "pdf")
                            {
                                string tempTxt = Path.GetTempFileName() + ".txt";
                                ExportHelper.ExportToTxt(data, currentDate, tempTxt, chkIncludeUnsubmitted.Checked);
                                ExportHelper.ExportToPdf(tempTxt, filePath);
                                File.Delete(tempTxt);
                            }
                            else if (format == "jpg")
                                ExportHelper.ExportToJpg(this, filePath);
                            else if (format == "html")
                                ExportHelper.ExportToHtml(data, currentDate, filePath, chkIncludeUnsubmitted.Checked);
                        }

                        MessageBox.Show($"导出成功：{filePath}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dlg.DialogResult = DialogResult.OK;
                        dlg.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };

                btnShare.Click += async (s, e) =>
                {
                    // 先导出再分享
                    string format = cmbFormat.SelectedItem.ToString();
                    string fileName = $"作业_{DateTime.Now:yyyyMMdd_HHmmss}.{format}";
                    string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                    try
                    {
                        if (chkRangeMode.Checked)
                        {
                            // 范围导出暂不支持分享（简单处理）
                            MessageBox.Show("暂不支持范围导出分享", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        var data = homeworkData;
                        if (format == "txt")
                            ExportHelper.ExportToTxt(data, currentDate, filePath, chkIncludeUnsubmitted.Checked);
                        else if (format == "pdf")
                        {
                            string tempTxt = Path.GetTempFileName() + ".txt";
                            ExportHelper.ExportToTxt(data, currentDate, tempTxt, chkIncludeUnsubmitted.Checked);
                            ExportHelper.ExportToPdf(tempTxt, filePath);
                            File.Delete(tempTxt);
                        }
                        else if (format == "jpg")
                            ExportHelper.ExportToJpg(this, filePath);
                        else if (format == "html")
                            ExportHelper.ExportToHtml(data, currentDate, filePath, chkIncludeUnsubmitted.Checked);

                        await ShareHelper.ShowShareUIAsync(filePath);
                        dlg.DialogResult = DialogResult.OK;
                        dlg.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"分享失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };

                btnCancel.Click += (s, e) => dlg.Close();

                dlg.Controls.Add(chkIncludeUnsubmitted);
                dlg.Controls.Add(chkRangeMode);
                dlg.Controls.Add(dtpStart);
                dlg.Controls.Add(lblTo);
                dlg.Controls.Add(dtpEnd);
                dlg.Controls.Add(cmbFormat);
                dlg.Controls.Add(btnExport);
                dlg.Controls.Add(btnShare);
                dlg.Controls.Add(btnCancel);

                dlg.ShowDialog(this);
            }
        }

        // ---------- 内联编辑 ----------
        private void StartInlineEdit(int subjectIndex, EditFieldType fieldType, Rectangle fieldRect)
        {
            if (editingSubjectIndex != -1)
                FinishInlineEdit();

            editingSubjectIndex = subjectIndex;
            currentEditType = fieldType;

            string subject = currentSubjects[subjectIndex];
            Point screenLoc = MapToScreen(fieldRect.Location);
            Size screenSize = new Size(
                (int)(fieldRect.Width * scaleFactor),
                (int)(fieldRect.Height * scaleFactor)
            );

            string currentText = "";
            if (fieldType == EditFieldType.Subject)
                currentText = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
            else if (fieldType == EditFieldType.Unsubmitted)
                currentText = homeworkData.Unsubmitted.ContainsKey(subject) ? homeworkData.Unsubmitted[subject] : "";

            Font editFont = font30;
            if (fullscreen)
            {
                int newLevel = appConfig.FontSizeLevel + 1;
                if (newLevel > 2) newLevel = 2;
                float scale = fontScales[newLevel];
                editFont = new Font("微软雅黑", 20 * scale);
            }

            var textBox = new TextBox
            {
                Multiline = true,
                WordWrap = true,
                ScrollBars = ScrollBars.Vertical,
                Location = screenLoc,
                Size = screenSize,
                Text = currentText,
                Font = editFont,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(40, 40, 40),
                ForeColor = Color.White
            };
            textBox.LostFocus += InlineEdit_LostFocus;
            textBox.KeyDown += InlineEdit_KeyDown;
            inlineEditControl = textBox;

            this.Controls.Add(inlineEditControl);
            inlineEditControl.Focus();
        }

        private void InlineEdit_LostFocus(object sender, EventArgs e)
        {
            FinishInlineEdit();
        }

        private void InlineEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    e.SuppressKeyPress = true;
                    FinishInlineEdit();
                }
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                CancelInlineEdit();
            }
            else if (e.Control && e.KeyCode == Keys.I)
            {
                e.SuppressKeyPress = true;
                if (inlineEditControl is TextBox tb && appConfig.EnableMarkdown && currentEditType == EditFieldType.Subject)
                {
                    tb.LostFocus -= InlineEdit_LostFocus;
                    try
                    {
                        using (var dialog = new InsertFormulaDialog())
                        {
                            if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.FormulaText))
                            {
                                string unicodeFormula = MarkdownRenderer.ConvertLatexToUnicode(dialog.FormulaText);
                                int start = tb.SelectionStart;
                                int length = tb.SelectionLength;
                                if (length > 0)
                                {
                                    tb.Text = tb.Text.Remove(start, length).Insert(start, unicodeFormula);
                                    tb.SelectionStart = start + unicodeFormula.Length;
                                }
                                else
                                {
                                    tb.Text = tb.Text + unicodeFormula;
                                    tb.SelectionStart = tb.Text.Length;
                                }
                                tb.ScrollToCaret();
                            }
                        }
                    }
                    finally
                    {
                        tb.LostFocus += InlineEdit_LostFocus;
                        tb.Focus();
                    }
                }
            }
        }

        private void FinishInlineEdit()
        {
            if (editingSubjectIndex == -1 || inlineEditControl == null) return;

            TextBox textBox = inlineEditControl as TextBox;
            textBox.LostFocus -= InlineEdit_LostFocus;
            textBox.KeyDown -= InlineEdit_KeyDown;

            string subject = currentSubjects[editingSubjectIndex];
            string newText = textBox.Text;

            if (currentEditType == EditFieldType.Subject)
                homeworkData.Subjects[subject] = newText;
            else if (currentEditType == EditFieldType.Unsubmitted)
                homeworkData.Unsubmitted[subject] = newText;

            SaveHomeworkData();

            this.Controls.Remove(inlineEditControl);
            inlineEditControl.Dispose();
            inlineEditControl = null;
            editingSubjectIndex = -1;
            currentEditType = EditFieldType.None;
            if (!this.IsDisposed)
                Invalidate();
        }

        private void CancelInlineEdit()
        {
            if (editingSubjectIndex == -1 || inlineEditControl == null) return;

            TextBox textBox = inlineEditControl as TextBox;
            textBox.LostFocus -= InlineEdit_LostFocus;
            textBox.KeyDown -= InlineEdit_KeyDown;

            this.Controls.Remove(inlineEditControl);
            inlineEditControl.Dispose();
            inlineEditControl = null;
            editingSubjectIndex = -1;
            currentEditType = EditFieldType.None;
            if (!this.IsDisposed)
                Invalidate();
        }

        // ---------- 轮播模式 ----------
        private void RotateManual(int direction)
        {
            List<int> nonEmpty = new List<int>();
            for (int i = 0; i < currentSubjects.Length; i++)
            {
                string subj = currentSubjects[i];
                if (homeworkData.Subjects.ContainsKey(subj) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[subj]))
                    nonEmpty.Add(i);
            }
            if (nonEmpty.Count == 0)
            {
                rotationMode = false;
                rotationTimer.Stop();
                MessageBox.Show("所有科目都没有作业内容，已退出轮播模式", "提示");
                return;
            }
            int curIdx = nonEmpty.IndexOf(rotationIndex);
            if (curIdx < 0) curIdx = 0;
            curIdx = (curIdx + direction + nonEmpty.Count) % nonEmpty.Count;
            rotationIndex = nonEmpty[curIdx];
            Invalidate();
        }

        private void RotateNext()
        {
            RotateManual(1);
        }

        // ---------- 图片加载 ----------
        private void LoadImages()
        {
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
            buttonImage = LoadImage(Path.Combine(imagePath, "按钮.png"), new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));
            historyBtnImage = LoadImage(Path.Combine(imagePath, "更多.png"), new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));

            Image originalBack = LoadImage(Path.Combine(imagePath, "返回.png"), originalSize: true);
            if (originalBack != null)
            {
                int targetHeight = BTN_SQUARE_SIZE;
                int targetWidth = (int)((float)originalBack.Width / originalBack.Height * targetHeight);
                backBtnImage = new Bitmap(originalBack, new Size(targetWidth, targetHeight));
            }

            leftArrowImage = LoadImage(Path.Combine(imagePath, "箭头图片.png"), new Size(50, 50));
            if (leftArrowImage != null)
            {
                rightArrowImage = (Image)new Bitmap(leftArrowImage);
                rightArrowImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            }

            // 图标列表（注意：新增“导出”按钮）
            string[] iconNames = { "编辑", "返回", "历史", "轮换", "全屏", "设置", "缩小", "完成", "未交人员", "关闭", "最小化", "展开", "收起", "导出" };
            foreach (string name in iconNames)
            {
                string filePath = Path.Combine(imagePath, name + ".png");
                if (File.Exists(filePath))
                {
                    // 直接加载原始图片，不再缩放
                    buttonIcons[name] = Image.FromFile(filePath);
                }
            }
        }

        private Image LoadImage(string path, bool originalSize)
        {
            try
            {
                if (File.Exists(path))
                    return Image.FromFile(path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载图片失败: {path}, 错误: {ex.Message}");
            }
            return null;
        }

        private Image LoadImage(string path, Size size)
        {
            try
            {
                if (File.Exists(path))
                {
                    Image img = Image.FromFile(path);
                    return new Bitmap(img, size);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载图片失败: {path}, 错误: {ex.Message}");
            }
            return null;
        }

        // ---------- 颜色解析辅助 ----------
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

        // ---------- 获取提交时间画刷（根据值） ----------
        private Brush GetDueTimeBrush(string value)
        {
            if (value.StartsWith("晚修"))
            {
                if (int.TryParse(value.Substring(2), out int index))
                {
                    index--;
                    switch (index)
                    {
                        case 0: return RED_SEMI;
                        case 1: return ORANGE_SEMI;
                        case 2: return GREEN_SEMI;
                        case 3: return BLUE_SEMI;
                        case 4: return PURPLE_SEMI;
                        case 5: return DARKORANGE_SEMI;
                        default: return new SolidBrush(TEXT_COLOR);
                    }
                }
            }
            return new SolidBrush(TEXT_COLOR);
        }

        // ---------- 绘制 ----------
        protected override void OnPaint(PaintEventArgs e)
        {
             base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;  // 新增

            g.TranslateTransform(offset.X, offset.Y);
            g.ScaleTransform(scaleFactor, scaleFactor);

            if (!_micaEnabled)
            {
                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(_backgroundAlpha, 32, 32, 32)))
                {
                    g.FillRectangle(bgBrush, new Rectangle(0, 0, VIRTUAL_SIZE.Width, VIRTUAL_SIZE.Height));
                }
            }

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
            using (SolidBrush barBrush = new SolidBrush(Color.FromArgb(barAlpha, barColor.R, barColor.G, barColor.B)))
            using (GraphicsPath barPath = CreateRoundedRectPath(barRect, ROUND_RADIUS, bottomOnly: true))
            {
                g.FillPath(barBrush, barPath);
                using (Pen borderPen = new Pen(Color.FromArgb(120, 255, 255, 255), 1))
                {
                    g.DrawPath(borderPen, barPath);
                }
            }

            DrawDateInfo(g, barRect);
            if (rotationMode)
                DrawRotationView(g);
            else if (unsubmittedMode)
                DrawUnsubmittedView(g);
            else
                DrawGrid(g);
            DrawButtons(g, barRect);

            // 编辑模式下显示提示
            if (editMode && !rotationMode && !unsubmittedMode && appConfig.EnableMarkdown)
            {
                using (Font hintFont = new Font("微软雅黑", 9, FontStyle.Italic))
                using (Brush hintBrush = new SolidBrush(Color.FromArgb(180, Color.Gray)))
                {
                    string hintText = "提示：按 Ctrl + I 可插入数学公式";
                    SizeF size = g.MeasureString(hintText, hintFont);
                    float x = (VIRTUAL_SIZE.Width - size.Width) / 2;
                    float y = VIRTUAL_SIZE.Height - 25;
                    g.DrawString(hintText, hintFont, hintBrush, x, y);
                }
            }

            g.ResetTransform();
        }

        private void DrawDateInfo(Graphics g, Rectangle barRect)
        {
            DateTime now = historyMode && historyDate.HasValue ? historyDate.Value : currentDate;
            string[] weekdays = { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日" };
            string weekdayText = weekdays[(int)now.DayOfWeek == 0 ? 6 : (int)now.DayOfWeek - 1];
            string dateText = now.ToString("yyyy年MM月dd日");
            if (historyMode) dateText += " (历史作业)";

            using (Brush whiteBrush = new SolidBrush(Color.White))
            {
                if (rotationMode || unsubmittedMode)
                {
                    int totalHeight = font20.Height + font12.Height;
                    int startY = barRect.Top + (barRect.Height - totalHeight) / 2;

                    SizeF weekdaySize = g.MeasureString(weekdayText, font20);
                    g.DrawString(weekdayText, font20, whiteBrush, new PointF((VIRTUAL_SIZE.Width - weekdaySize.Width) / 2, startY));

                    SizeF dateSize = g.MeasureString(dateText, font12);
                    g.DrawString(dateText, font12, whiteBrush, new PointF((VIRTUAL_SIZE.Width - dateSize.Width) / 2, startY + font20.Height));
                }
                else
                {
                    int totalHeight = font20.Height + font12.Height;
                    int startY = barRect.Top + (barRect.Height - totalHeight) / 2;
                    int x = barRect.Left + BAR_PADDING;

                    g.DrawString(weekdayText, font20, whiteBrush, new PointF(x, startY));
                    g.DrawString(dateText, font12, whiteBrush, new PointF(x, startY + font20.Height));
                }
            }
        }

        private void DrawButtons(Graphics g, Rectangle barRect)
        {
            if (rotationMode || unsubmittedMode)
            {
                int backY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;
                backBtnRect = new Rectangle(35, backY, BTN_SQUARE_SIZE, BTN_SQUARE_SIZE);
                DrawTransparentButton(g, backBtnRect, "返回", buttonIcons.ContainsKey("返回") ? buttonIcons["返回"] : null, ref backBtnRect, backBtnRect.Location, _backPressed);

                int btnY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;
                int totalButtonsWidth = 4 * BTN_SQUARE_SIZE + 3 * 10; // 增加导出按钮
                int startX = barRect.Right - totalButtonsWidth - 20;

                Point editPos = new Point(startX, btnY);
                Point historyPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point fullscreenPos = new Point(historyPos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point exportPos = new Point(fullscreenPos.X + BTN_SQUARE_SIZE + 10, btnY);

                string editButtonText = unsubmittedMode ? (editMode ? "完成" : "登记") : (editMode ? "完成" : "编辑");
                string editIconKey = editMode ? "完成" : (unsubmittedMode ? "编辑" : "编辑");
                DrawTransparentButton(g, editBtnRect, editButtonText, buttonIcons.ContainsKey(editIconKey) ? buttonIcons[editIconKey] : null, ref editBtnRect, editPos, _editPressed);

                string historyIconKey = historyMode ? "返回" : "历史";
                DrawTransparentButton(g, historyBtnRect, historyMode ? "返回" : "历史", buttonIcons.ContainsKey(historyIconKey) ? buttonIcons[historyIconKey] : null, ref historyBtnRect, historyPos, _historyPressed);

                string fullscreenIconKey = fullscreen ? "缩小" : "全屏";
                DrawTransparentButton(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreenIconKey) ? buttonIcons[fullscreenIconKey] : null, ref fullscreenBtnRect, fullscreenPos, _fullscreenPressed);

                DrawTransparentButton(g, exportBtnRect, "导出", buttonIcons.ContainsKey("导出") ? buttonIcons["导出"] : null, ref exportBtnRect, exportPos, _exportPressed);

                if (unsubmittedMode)
                {
                    int totalPages = (currentSubjects.Length + UNSUBMITTED_ROWS_PER_PAGE - 1) / UNSUBMITTED_ROWS_PER_PAGE;
                    if (totalPages > 1)
                    {
                        leftArrowRect = new Rectangle(50, VIRTUAL_SIZE.Height / 2 - 25, 50, 50);
                        rightArrowRect = new Rectangle(VIRTUAL_SIZE.Width - 50 - 50, VIRTUAL_SIZE.Height / 2 - 25, 50, 50);
                        if (leftArrowImage != null)
                            g.DrawImage(leftArrowImage, leftArrowRect);
                        else
                        {
                            using (SolidBrush arrowBrush = new SolidBrush(Color.DarkGray))
                                g.FillPolygon(arrowBrush, new Point[] {
                                    new Point(leftArrowRect.Right, leftArrowRect.Top),
                                    new Point(leftArrowRect.Right, leftArrowRect.Bottom),
                                    new Point(leftArrowRect.Left, leftArrowRect.Top + leftArrowRect.Height/2)
                                });
                        }
                        if (rightArrowImage != null)
                            g.DrawImage(rightArrowImage, rightArrowRect);
                        else
                        {
                            using (SolidBrush arrowBrush = new SolidBrush(Color.DarkGray))
                                g.FillPolygon(arrowBrush, new Point[] {
                                    new Point(rightArrowRect.Left, rightArrowRect.Top),
                                    new Point(rightArrowRect.Left, rightArrowRect.Bottom),
                                    new Point(rightArrowRect.Right, rightArrowRect.Top + rightArrowRect.Height/2)
                                });
                        }
                    }
                }
                else if (rotationMode)
                {
                    leftArrowRect = new Rectangle(50, VIRTUAL_SIZE.Height / 2 - 25, 50, 50);
                    rightArrowRect = new Rectangle(VIRTUAL_SIZE.Width - 50 - 50, VIRTUAL_SIZE.Height / 2 - 25, 50, 50);
                    if (leftArrowImage != null)
                        g.DrawImage(leftArrowImage, leftArrowRect);
                    else
                    {
                        using (SolidBrush arrowBrush = new SolidBrush(Color.DarkGray))
                            g.FillPolygon(arrowBrush, new Point[] {
                                new Point(leftArrowRect.Right, leftArrowRect.Top),
                                new Point(leftArrowRect.Right, leftArrowRect.Bottom),
                                new Point(leftArrowRect.Left, leftArrowRect.Top + leftArrowRect.Height/2)
                            });
                    }
                    if (rightArrowImage != null)
                        g.DrawImage(rightArrowImage, rightArrowRect);
                    else
                    {
                        using (SolidBrush arrowBrush = new SolidBrush(Color.DarkGray))
                            g.FillPolygon(arrowBrush, new Point[] {
                                new Point(rightArrowRect.Left, rightArrowRect.Top),
                                new Point(rightArrowRect.Left, rightArrowRect.Bottom),
                                new Point(rightArrowRect.Right, rightArrowRect.Top + rightArrowRect.Height/2)
                            });
                    }
                }
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

                    DrawTransparentButton(g, expandBtnRect, "展开", buttonIcons.ContainsKey("展开") ? buttonIcons["展开"] : null, ref expandBtnRect, expandPos, _expandPressed);
                    DrawTransparentButton(g, editBtnRect, editMode ? "完成" : "编辑", buttonIcons.ContainsKey(editMode ? "完成" : "编辑") ? buttonIcons[editMode ? "完成" : "编辑"] : null, ref editBtnRect, editPos, _editPressed);

                    if (isWin10)
                        DrawTransparentButton(g, minimizeBtnRect, "最小化", buttonIcons.ContainsKey("最小化") ? buttonIcons["最小化"] : null, ref minimizeBtnRect, thirdPos, _minimizePressed);
                    else
                        DrawTransparentButton(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreen ? "缩小" : "全屏") ? buttonIcons[fullscreen ? "缩小" : "全屏"] : null, ref fullscreenBtnRect, thirdPos, _fullscreenPressed);
                }
                else
                {
                    if (isWin10)
                    {
                        int totalButtonsWidth = 9 * BTN_SQUARE_SIZE + 8 * 10;
                        int startX = barRect.Right - totalButtonsWidth - 20;

                        Point collapsePos = new Point(startX, btnY);
                        Point settingsPos = new Point(collapsePos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point rotatePos = new Point(settingsPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point unsubmittedPos = new Point(rotatePos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point exportPos = new Point(unsubmittedPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point editPos = new Point(exportPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point historyPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point minimizePos = new Point(historyPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point closePos = new Point(minimizePos.X + BTN_SQUARE_SIZE + 10, btnY);

                        DrawTransparentButton(g, collapseBtnRect, "收起", buttonIcons.ContainsKey("收起") ? buttonIcons["收起"] : null, ref collapseBtnRect, collapsePos, _collapsePressed);
                        DrawTransparentButton(g, settingsBtnRect, "设置", buttonIcons.ContainsKey("设置") ? buttonIcons["设置"] : null, ref settingsBtnRect, settingsPos, _settingsPressed);
                        DrawTransparentButton(g, rotateBtnRect, "轮换", buttonIcons.ContainsKey("轮换") ? buttonIcons["轮换"] : null, ref rotateBtnRect, rotatePos, _rotatePressed);
                        DrawTransparentButton(g, unsubmittedBtnRect, "未交", buttonIcons.ContainsKey("未交人员") ? buttonIcons["未交人员"] : null, ref unsubmittedBtnRect, unsubmittedPos, _unsubmittedPressed);
                        DrawTransparentButton(g, exportBtnRect, "导出", buttonIcons.ContainsKey("导出") ? buttonIcons["导出"] : null, ref exportBtnRect, exportPos, _exportPressed);
                        DrawTransparentButton(g, editBtnRect, editMode ? "完成" : "编辑", buttonIcons.ContainsKey(editMode ? "完成" : "编辑") ? buttonIcons[editMode ? "完成" : "编辑"] : null, ref editBtnRect, editPos, _editPressed);
                        DrawTransparentButton(g, historyBtnRect, historyMode ? "返回" : "历史", buttonIcons.ContainsKey(historyMode ? "返回" : "历史") ? buttonIcons[historyMode ? "返回" : "历史"] : null, ref historyBtnRect, historyPos, _historyPressed);
                        DrawTransparentButton(g, minimizeBtnRect, "最小化", buttonIcons.ContainsKey("最小化") ? buttonIcons["最小化"] : null, ref minimizeBtnRect, minimizePos, _minimizePressed);
                        DrawTransparentButton(g, closeBtnRect, "关闭", buttonIcons.ContainsKey("关闭") ? buttonIcons["关闭"] : null, ref closeBtnRect, closePos, _closePressed);
                    }
                    else
                    {
                        int totalButtonsWidth = 10 * BTN_SQUARE_SIZE + 9 * 10;
                        int startX = barRect.Right - totalButtonsWidth - 20;

                        Point collapsePos = new Point(startX, btnY);
                        Point settingsPos = new Point(collapsePos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point rotatePos = new Point(settingsPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point unsubmittedPos = new Point(rotatePos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point exportPos = new Point(unsubmittedPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point editPos = new Point(exportPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point historyPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point fullscreenPos = new Point(historyPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point minimizePos = new Point(fullscreenPos.X + BTN_SQUARE_SIZE + 10, btnY);
                        Point closePos = new Point(minimizePos.X + BTN_SQUARE_SIZE + 10, btnY);

                        DrawTransparentButton(g, collapseBtnRect, "收起", buttonIcons.ContainsKey("收起") ? buttonIcons["收起"] : null, ref collapseBtnRect, collapsePos, _collapsePressed);
                        DrawTransparentButton(g, settingsBtnRect, "设置", buttonIcons.ContainsKey("设置") ? buttonIcons["设置"] : null, ref settingsBtnRect, settingsPos, _settingsPressed);
                        DrawTransparentButton(g, rotateBtnRect, "轮换", buttonIcons.ContainsKey("轮换") ? buttonIcons["轮换"] : null, ref rotateBtnRect, rotatePos, _rotatePressed);
                        DrawTransparentButton(g, unsubmittedBtnRect, "未交", buttonIcons.ContainsKey("未交人员") ? buttonIcons["未交人员"] : null, ref unsubmittedBtnRect, unsubmittedPos, _unsubmittedPressed);
                        DrawTransparentButton(g, exportBtnRect, "导出", buttonIcons.ContainsKey("导出") ? buttonIcons["导出"] : null, ref exportBtnRect, exportPos, _exportPressed);
                        DrawTransparentButton(g, editBtnRect, editMode ? "完成" : "编辑", buttonIcons.ContainsKey(editMode ? "完成" : "编辑") ? buttonIcons[editMode ? "完成" : "编辑"] : null, ref editBtnRect, editPos, _editPressed);
                        DrawTransparentButton(g, historyBtnRect, historyMode ? "返回" : "历史", buttonIcons.ContainsKey(historyMode ? "返回" : "历史") ? buttonIcons[historyMode ? "返回" : "历史"] : null, ref historyBtnRect, historyPos, _historyPressed);
                        DrawTransparentButton(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreen ? "缩小" : "全屏") ? buttonIcons[fullscreen ? "缩小" : "全屏"] : null, ref fullscreenBtnRect, fullscreenPos, _fullscreenPressed);
                        DrawTransparentButton(g, minimizeBtnRect, "最小化", buttonIcons.ContainsKey("最小化") ? buttonIcons["最小化"] : null, ref minimizeBtnRect, minimizePos, _minimizePressed);
                        DrawTransparentButton(g, closeBtnRect, "关闭", buttonIcons.ContainsKey("关闭") ? buttonIcons["关闭"] : null, ref closeBtnRect, closePos, _closePressed);
                    }
                }
            }
        }

        private void DrawTransparentButton(Graphics g, Rectangle rect, string text, Image icon, ref Rectangle targetRect, Point pos, bool pressed)
        {
            targetRect = new Rectangle(pos, new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));

            const int iconTopMargin = 3;
            const float textBottomMargin = 1;

            if (icon != null)
            {
                // 图标高度占按钮高度的 50%
                int targetHeight = (int)(targetRect.Height * 0.5);
                int targetWidth = (int)((float)icon.Width / icon.Height * targetHeight);
                
                // 水平居中
                int iconX = targetRect.Left + (targetRect.Width - targetWidth) / 2;
                // 垂直方向：原本垂直居中后再向上移动 5 像素（可调整）
                int iconY = targetRect.Top + (targetRect.Height - targetHeight) / 2 - 7;
                
                // 设置高质量插值模式
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(icon, iconX, iconY, targetWidth, targetHeight);
            }

            using (Font smallFont = new Font("微软雅黑", 8))
            {
                SizeF textSize = g.MeasureString(text, smallFont);
                float textX = targetRect.Left + (targetRect.Width - textSize.Width) / 2;
                float textY = targetRect.Bottom - textSize.Height - textBottomMargin;
                g.DrawString(text, smallFont, Brushes.White, textX, textY);
            }
        }

        private StringFormat CenterStringFormat()
        {
            return new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        }

        private void DrawGrid(Graphics g)
        {
            float opacityFactor = appConfig.CardOpacity / 100f;

            for (int i = 0; i < gridRects.Count; i++)
            {
                Rectangle rect = gridRects[i];

                int shadowOffset = 3;
                Rectangle shadowRect = new Rectangle(rect.X + shadowOffset, rect.Y + shadowOffset, rect.Width, rect.Height);
                using (GraphicsPath shadowPath = CreateRoundedRectPath(shadowRect, ROUND_RADIUS))
                {
                    using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(30, 0, 0, 0)))
                    {
                        g.FillPath(shadowBrush, shadowPath);
                    }
                }

                int topAlpha = (int)(220 * opacityFactor);
                int bottomAlpha = (int)(160 * opacityFactor);
                using (GraphicsPath path = CreateRoundedRectPath(rect, ROUND_RADIUS))
                {
                    using (LinearGradientBrush bgBrush = new LinearGradientBrush(
                        rect,
                        Color.FromArgb(topAlpha, 255, 255, 255),
                        Color.FromArgb(bottomAlpha, 240, 240, 255),
                        LinearGradientMode.Vertical))
                    {
                        g.FillPath(bgBrush, path);
                    }

                    string subject = currentSubjects[i];
                    string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";

                    bool highlightActive = _activeEvenings.Contains(dueTime);
                    bool highlightFlash = _flashingEvenings.Contains(dueTime);
                    bool highlightGray = _grayEvenings.Contains(dueTime);
                    Color borderColor = Color.FromArgb(100, 128, 128, 128);
                    float borderWidth = 1;

                    if (_debugFlashing || highlightFlash)
                    {
                        int alpha = flashStep;
                        using (SolidBrush redBrush = new SolidBrush(Color.FromArgb(alpha, 255, 0, 0)))
                        {
                            g.FillPath(redBrush, path);
                        }
                        DrawLaserBorder(g, path, rect);
                        borderColor = Color.FromArgb(200, 255, 0, 0);
                    }
                    else if (highlightActive)
                    {
                        borderColor = SystemColors.Highlight;
                        borderWidth = 2;
                    }

                    using (Pen borderPen = new Pen(borderColor, borderWidth))
                    {
                        g.DrawPath(borderPen, path);
                    }
                }

                if (i < currentSubjects.Length)
                {
                    string subject = currentSubjects[i];

                    int subjectY = rect.Top + 50 - font22.Height;
                    g.DrawString(subject, font22, new SolidBrush(TEXT_COLOR), new PointF(rect.Left + 10, subjectY));

                    int lineY = rect.Top + 50;
                    g.DrawLine(Pens.Gray, rect.Left + 10, lineY, rect.Right - 10, lineY);

                    if (!EditMode)
                    {
                        string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
                        string prefix = "提交时间：";
                        string displayTime = string.IsNullOrEmpty(dueTime) ? "无" : dueTime;

                        float prefixWidth = g.MeasureString(prefix, fontSmall).Width;
                        float timeWidth = g.MeasureString(displayTime, font22).Width;
                        float totalWidth = prefixWidth + timeWidth;
                        int rightX = rect.Right - 10;
                        float startX = rightX - totalWidth;
                        int timeY = lineY - font22.Height;
                        int prefixY = lineY - fontSmall.Height;

                        g.DrawString(prefix, fontSmall, new SolidBrush(TEXT_COLOR), startX, prefixY);
                        Brush timeBrush = GetDueTimeBrush(displayTime);
                        g.DrawString(displayTime, font22, timeBrush, startX + prefixWidth, timeY);
                        if (timeBrush != RED_SEMI && timeBrush != ORANGE_SEMI && timeBrush != GREEN_SEMI &&
                            timeBrush != BLUE_SEMI && timeBrush != PURPLE_SEMI && timeBrush != DARKORANGE_SEMI)
                            timeBrush.Dispose();
                    }

                    Rectangle textArea = new Rectangle(
                        rect.Left + 10,
                        rect.Top + 60,
                        rect.Width - 20,
                        rect.Height - 70);

                    if (editingSubjectIndex == i && currentEditType == EditFieldType.Subject)
                    {
                        // 正在编辑，不绘制
                    }
                    else if (homeworkData.Subjects.ContainsKey(subject) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[subject]))
                    {
                        float offset = scrollOffsets.ContainsKey(i) ? scrollOffsets[i] : 0f;
                        DrawTextInArea(g, homeworkData.Subjects[subject], textArea, font30, false, offset);
                    }
                    else
                    {
                        int emptyAlpha = (int)(255 * opacityFactor);
                        using (GraphicsPath textAreaPath = CreateRoundedRectPath(textArea, ROUND_RADIUS))
                        {
                            using (SolidBrush lightBrush = new SolidBrush(Color.FromArgb(emptyAlpha, 255, 255, 255)))
                            {
                                g.FillPath(lightBrush, textAreaPath);
                            }
                            using (Pen lightPen = new Pen(Color.FromArgb(120, 128, 128, 128)))
                            {
                                g.DrawPath(lightPen, textAreaPath);
                            }
                        }
                        string hint = editMode ? "点我编辑作业" : "今天暂时没有此项作业";
                        g.DrawString(hint, hintFont, RED_SEMI, textArea, CenterStringFormat());
                    }
                }
            }
        }

        private void DrawRotationView(Graphics g)
        {
            float opacityFactor = appConfig.CardOpacity / 100f;
            int topAlpha = (int)(220 * opacityFactor);
            int bottomAlpha = (int)(160 * opacityFactor);

            Rectangle bigRect = new Rectangle(150, 100, 900, 475);
            using (GraphicsPath path = CreateRoundedRectPath(bigRect, ROUND_RADIUS))
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(
                    bigRect,
                    Color.FromArgb(topAlpha, 255, 255, 255),
                    Color.FromArgb(bottomAlpha, 240, 240, 255),
                    LinearGradientMode.Vertical))
                {
                    g.FillPath(brush, path);
                }

                string subject = currentSubjects[rotationIndex];
                string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";

                bool highlightActive = _activeEvenings.Contains(dueTime);
                bool highlightFlash = _flashingEvenings.Contains(dueTime);
                bool highlightGray = _grayEvenings.Contains(dueTime);
                Color borderColor = Color.FromArgb(100, 128, 128, 128);
                float borderWidth = 1;

                if (_debugFlashing || highlightFlash)
                {
                    int alpha = flashStep;
                    using (SolidBrush redBrush = new SolidBrush(Color.FromArgb(alpha, 255, 0, 0)))
                    {
                        g.FillPath(redBrush, path);
                    }
                    DrawLaserBorder(g, path, bigRect);
                    borderColor = Color.FromArgb(200, 255, 0, 0);
                }
                else if (highlightActive)
                {
                    borderColor = SystemColors.Highlight;
                    borderWidth = 2;
                }

                using (Pen borderPen = new Pen(borderColor, borderWidth))
                {
                    g.DrawPath(borderPen, path);
                }
            }

            string subject2 = currentSubjects[rotationIndex];
            g.DrawString(subject2, font36, Brushes.White, new PointF(600 - (int)g.MeasureString(subject2, font36).Width / 2, bigRect.Top + 30));

            int lineY = bigRect.Top + 120;
            g.DrawLine(new Pen(Color.Gray, 2), bigRect.Left + 50, lineY, bigRect.Right - 50, lineY);

            string dueTime2 = homeworkData.DueTimes.ContainsKey(subject2) ? homeworkData.DueTimes[subject2] : "";
            string displayTime;
            if (string.IsNullOrEmpty(dueTime2))
                displayTime = "无";
            else
                displayTime = dueTime2;

            string prefix = "提交时间：";
            float prefixWidth = g.MeasureString(prefix, fontSmall).Width;
            float timeWidth = g.MeasureString(displayTime, font22).Width;
            float totalWidth = prefixWidth + timeWidth;
            int rightX = bigRect.Right - 60;
            float startX = rightX - totalWidth;

            int prefixY = lineY - fontSmall.Height;
            int timeY = lineY - font22.Height;

            g.DrawString(prefix, fontSmall, new SolidBrush(TEXT_COLOR), startX, prefixY);
            Brush timeBrush = GetDueTimeBrush(displayTime);
            g.DrawString(displayTime, font22, timeBrush, startX + prefixWidth, timeY);
            if (timeBrush != RED_SEMI && timeBrush != ORANGE_SEMI && timeBrush != GREEN_SEMI &&
                timeBrush != BLUE_SEMI && timeBrush != PURPLE_SEMI && timeBrush != DARKORANGE_SEMI)
                timeBrush.Dispose();

            Rectangle textArea = new Rectangle(bigRect.Left + 50, bigRect.Top + 150, bigRect.Width - 100, bigRect.Height - 200);
            if (homeworkData.Subjects.ContainsKey(subject2) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[subject2]))
            {
                float offset = scrollOffsets.ContainsKey(rotationIndex) ? scrollOffsets[rotationIndex] : 0f;
                DrawTextInArea(g, homeworkData.Subjects[subject2], textArea, font30, true, offset);
            }
            else
            {
                g.DrawString("今天暂时没有此项作业", font30, RED_SEMI, textArea, CenterStringFormat());
            }
        }

        private void DrawUnsubmittedView(Graphics g)
        {
            if (font30 == null || fontSmall == null || font22 == null || hintFont == null)
            {
                ApplyFontSettings();
                if (font30 == null) return;
            }

            float opacityFactor = appConfig.CardOpacity / 100f;
            int topAlpha = (int)(220 * opacityFactor);
            int bottomAlpha = (int)(160 * opacityFactor);

            Rectangle bigRect = new Rectangle(150, 100, 900, 475);
            using (GraphicsPath path = CreateRoundedRectPath(bigRect, ROUND_RADIUS))
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(
                    bigRect,
                    Color.FromArgb(topAlpha, 255, 255, 255),
                    Color.FromArgb(bottomAlpha, 240, 240, 255),
                    LinearGradientMode.Vertical))
                {
                    g.FillPath(brush, path);
                }

                int totalSubjects = currentSubjects.Length;
                int totalPages = (totalSubjects + UNSUBMITTED_ROWS_PER_PAGE - 1) / UNSUBMITTED_ROWS_PER_PAGE;
                int startIndex = unsubmittedPage * UNSUBMITTED_ROWS_PER_PAGE;
                int endIndex = Math.Min(startIndex + UNSUBMITTED_ROWS_PER_PAGE, totalSubjects);

                string title = $"未交名单 (第 {unsubmittedPage + 1}/{totalPages} 页)";
                g.DrawString(title, font30, Brushes.White, new PointF(bigRect.Left + 50, bigRect.Top + 30));

                int lineY = bigRect.Top + 80;
                g.DrawLine(new Pen(Color.Gray, 2), bigRect.Left + 50, lineY, bigRect.Right - 50, lineY);

                int rowHeight = (bigRect.Height - 120) / UNSUBMITTED_ROWS_PER_PAGE;
                int baseY = bigRect.Top + 100;

                const int subjectToCountSpacing = 2;
                const int countToListSpacing = 5;
                const int listVerticalPadding = 3;

                for (int i = startIndex; i < endIndex; i++)
                {
                    string subject = currentSubjects[i];
                    int rowY = baseY + (i - startIndex) * rowHeight;

                    float subjectY = rowY + 5;
                    g.DrawString(subject, font22, new SolidBrush(TEXT_COLOR), new PointF(bigRect.Left + 50, subjectY));

                    string unsubmittedText = homeworkData.Unsubmitted.ContainsKey(subject) ? homeworkData.Unsubmitted[subject] : "";
                    int peopleCount = 0;
                    if (!string.IsNullOrWhiteSpace(unsubmittedText))
                    {
                        string pattern = @"[，,。；;、\s]+";
                        string[] parts = System.Text.RegularExpressions.Regex.Split(unsubmittedText, pattern);
                        peopleCount = parts.Count(p => !string.IsNullOrWhiteSpace(p));
                        if (peopleCount == 0 && !string.IsNullOrWhiteSpace(unsubmittedText))
                            peopleCount = 1;
                    }

                    string countText = peopleCount.ToString();
                    string suffixText = "人未交";
                    float countWidth = g.MeasureString(countText, font30).Width;
                    float statsX = bigRect.Left + 50;
                    float statsY = subjectY + font22.Height + subjectToCountSpacing;
                    g.DrawString(countText, font30, Brushes.Red, statsX, statsY);
                    g.DrawString(suffixText, fontSmall, Brushes.Red, statsX + countWidth, statsY + font30.Height - fontSmall.Height - 2);

                    float usedHeight = (subjectY - rowY) + font22.Height + subjectToCountSpacing + font30.Height;
                    float availableForList = rowHeight - usedHeight - countToListSpacing;
                    if (availableForList < listVerticalPadding * 2) availableForList = listVerticalPadding * 2;

                    float listAreaHeight = availableForList - listVerticalPadding * 2;
                    float listAreaTop = rowY + usedHeight + countToListSpacing + listVerticalPadding - 40;

                    RectangleF textRect = new RectangleF(
                        bigRect.Left + 200,
                        listAreaTop,
                        bigRect.Width - 250,
                        listAreaHeight);

                    if (editingSubjectIndex == i && currentEditType == EditFieldType.Unsubmitted)
                    {
                        // 正在编辑，不绘制
                    }
                    else if (!string.IsNullOrWhiteSpace(unsubmittedText))
                    {
                        GraphicsState state = g.Save();
                        g.SetClip(textRect);

                        List<string> lines = new List<string>();
                        string[] paragraphs = unsubmittedText.Split('\n');
                        foreach (string para in paragraphs)
                        {
                            if (string.IsNullOrEmpty(para)) continue;
                            string[] words = para.Split(' ');
                            string line = "";
                            foreach (string word in words)
                            {
                                string test = line + (line == "" ? "" : " ") + word;
                                if (g.MeasureString(test, font30).Width <= textRect.Width)
                                {
                                    line = test;
                                }
                                else
                                {
                                    if (line != "") lines.Add(line);
                                    line = word;
                                }
                            }
                            if (line != "") lines.Add(line);
                        }

                        float lineHeight = font30.GetHeight(g);
                        float totalTextHeight = lines.Count * lineHeight;
                        float verticalOffset = (textRect.Height - totalTextHeight) / 2;
                        if (verticalOffset < 0) verticalOffset = 0;

                        for (int j = 0; j < lines.Count; j++)
                        {
                            float y = textRect.Top + j * lineHeight + verticalOffset;
                            SizeF sz = g.MeasureString(lines[j], font30);
                            float x = textRect.Left + (textRect.Width - sz.Width) / 2;
                            g.DrawString(lines[j], font30, new SolidBrush(TEXT_COLOR), x, y);
                        }

                        g.Restore(state);
                    }
                    else
                    {
                        string hint = editMode ? "点此编辑名单" : "无未交人员";
                        SizeF hintSize = g.MeasureString(hint, hintFont);
                        float hintX = textRect.Left + (textRect.Width - hintSize.Width) / 2;
                        float hintY = textRect.Top + (textRect.Height - hintSize.Height) / 2;
                        g.DrawString(hint, hintFont, RED_SEMI, hintX, hintY);
                    }

                    if (i < endIndex - 1)
                        g.DrawLine(Pens.LightGray, bigRect.Left + 50, rowY + rowHeight, bigRect.Right - 50, rowY + rowHeight);
                }

                using (Pen borderPen = new Pen(Color.FromArgb(100, 128, 128, 128), 1))
                {
                    g.DrawPath(borderPen, path);
                }
            }
        }

        private void DrawLaserBorder(Graphics g, GraphicsPath path, Rectangle rect)
        {
            using (Pen laserPen = new Pen(Color.Red, 3))
            {
                laserPen.DashStyle = DashStyle.Custom;
                laserPen.DashPattern = new float[] { 8, 8 };
                laserPen.DashOffset = _laserOffset * 16;
                g.DrawPath(laserPen, path);
            }
        }

        // ---------- 文本绘制（支持 Markdown 渲染和 LaTeX 转换） ----------
        private void DrawTextInArea(Graphics g, string text, Rectangle area, Font font, bool center = false, float scrollOffset = 0f)
        {
            if (editMode || !appConfig.EnableMarkdown)
            {
                DrawPlainTextInArea(g, text, area, font, center, scrollOffset);
                return;
            }

            var paragraphs = MarkdownRenderer.ParseMarkdown(text);
            if (paragraphs.Count == 0) return;

            _markdownTotalHeight = 0;
            foreach (var para in paragraphs)
                _markdownTotalHeight += GetParagraphHeight(para, font);

            float y = area.Top - _markdownScrollOffset;
            float leftMargin = area.Left + 5;

            GraphicsState state = g.Save();
            g.SetClip(area);

            using (var defaultBrush = new SolidBrush(TEXT_COLOR))
            {
                foreach (var para in paragraphs)
                {
                    if (para.FormattedParts.Any(p => p.IsLatex))
                    {
                        string combinedText = string.Concat(para.FormattedParts.Select(p => p.Text));
                        float paraHeight = GetParagraphHeight(para, font);
                        Rectangle textRect = new Rectangle(area.Left, (int)y, area.Width, (int)paraHeight);
                        DrawPlainTextInArea(g, combinedText, textRect, font, center, 0);
                        y += paraHeight * 1.2f;
                        continue;
                    }

                    Font paraFont = font;
                    Brush paraBrush = defaultBrush;
                    float currentParaHeight = GetParagraphHeight(para, font);

                    if (para.Type == ParagraphType.Code)
                    {
                        paraFont = new Font("Consolas", font.Size, FontStyle.Regular);
                        paraBrush = new SolidBrush(Color.LightGray);
                        using (var bgBrush = new SolidBrush(Color.FromArgb(40, 40, 40)))
                        {
                            g.FillRectangle(bgBrush, leftMargin - 2, y, area.Width - 10, currentParaHeight);
                        }
                    }
                    else if (para.Type == ParagraphType.Quote)
                    {
                        paraFont = new Font(font.FontFamily, font.Size, FontStyle.Italic);
                        paraBrush = new SolidBrush(Color.Gray);
                        using (var lineBrush = new SolidBrush(Color.Gray))
                        {
                            g.FillRectangle(lineBrush, leftMargin - 2, y, 3, currentParaHeight);
                        }
                    }
                    else if (para.Type == ParagraphType.Heading)
                    {
                        float scale = 1.0f + (6 - para.HeadingLevel) * 0.2f;
                        paraFont = new Font(font.FontFamily, font.Size * scale, FontStyle.Bold);
                        currentParaHeight = paraFont.GetHeight(g);
                    }

                    if (para.IsListItem)
                    {
                        string prefix = para.IsOrderedList ? $"{para.ListItemNumber}. " : "• ";
                        DrawText(g, prefix, paraFont, defaultBrush, leftMargin, y);
                        leftMargin += paraFont.GetHeight(g) * 0.8f;
                    }

                    foreach (var part in para.FormattedParts)
                    {
                        Font partFont = new Font(paraFont.FontFamily, paraFont.Size, part.Style);
                        // 关键：优先使用 part.Color，否则根据段落类型决定颜色
                        Color textColor = part.Color ?? (para.Type == ParagraphType.Code ? Color.LightGray : TEXT_COLOR);
                        using (var brush = new SolidBrush(textColor))
                        {
                            SizeF size = g.MeasureString(part.Text, partFont);
                            g.DrawString(part.Text, partFont, brush, leftMargin, y);
                            leftMargin += size.Width;
                        }
                        partFont.Dispose();
                    }

                    y += currentParaHeight * 1.2f;
                    leftMargin = area.Left + 5;

                    if (para.Type == ParagraphType.Code || para.Type == ParagraphType.Quote || para.Type == ParagraphType.Heading)
                        paraFont?.Dispose();
                    if (paraBrush != defaultBrush) paraBrush?.Dispose();
                }
            }

            g.Restore(state);
        }

        private void DrawPlainTextInArea(Graphics g, string text, Rectangle area, Font font, bool center = false, float scrollOffset = 0f)
        {
            List<string> lines = new List<string>();
            string[] paragraphs = text.Split('\n');
            foreach (string para in paragraphs)
            {
                if (string.IsNullOrEmpty(para)) continue;
                string[] words = para.Split(' ');
                string line = "";
                foreach (string word in words)
                {
                    string test = line + (line == "" ? "" : " ") + word;
                    if (g.MeasureString(test, font).Width <= area.Width)
                    {
                        line = test;
                    }
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
                        else
                        {
                            line = word;
                        }
                    }
                }
                if (!string.IsNullOrEmpty(line)) lines.Add(line);
            }

            float lineHeight = font.GetHeight(g);

            GraphicsState state = g.Save();
            g.SetClip(area);

            using (SolidBrush textBrush = new SolidBrush(TEXT_COLOR))
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    float y = area.Top + i * lineHeight - scrollOffset;
                    if (center)
                    {
                        SizeF sz = g.MeasureString(lines[i], font);
                        g.DrawString(lines[i], font, textBrush, new PointF(area.Left + (area.Width - sz.Width) / 2, y));
                    }
                    else
                    {
                        g.DrawString(lines[i], font, textBrush, new PointF(area.Left + 5, y));
                    }
                }
            }

            g.Restore(state);
        }

        private void DrawText(Graphics g, string text, Font font, Brush brush, float x, float y)
        {
            g.DrawString(text, font, brush, x, y);
        }

        private float GetParagraphHeight(Paragraph para, Font baseFont)
        {
            if (para.Type == ParagraphType.Code)
            {
                int lineCount = 0;
                foreach (var part in para.FormattedParts)
                {
                    lineCount += part.Text.Split('\n').Length;
                }
                using (var codeFont = new Font("Consolas", baseFont.Size, FontStyle.Regular))
                {
                    return codeFont.GetHeight() * lineCount;
                }
            }
            else if (para.Type == ParagraphType.Quote)
            {
                using (var quoteFont = new Font(baseFont.FontFamily, baseFont.Size, FontStyle.Italic))
                {
                    return quoteFont.GetHeight();
                }
            }
            else if (para.Type == ParagraphType.Heading)
            {
                float scale = 1.0f + (6 - para.HeadingLevel) * 0.2f;
                using (var headingFont = new Font(baseFont.FontFamily, baseFont.Size * scale, FontStyle.Bold))
                {
                    return headingFont.GetHeight();
                }
            }
            else
            {
                return baseFont.GetHeight();
            }
        }

        // ---------- 时间提醒 ----------
        private void CheckEveningClassStates()
        {
            if (appConfig.EveningClassTimes == null || appConfig.EveningClassTimes.Count == 0)
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
            List<string> newActive = new List<string>();
            List<string> newFlashing = new List<string>();
            List<string> newGray = new List<string>();
            bool newFlashingExists = false;

            for (int i = 0; i < appConfig.EveningClassTimes.Count; i++)
            {
                var time = appConfig.EveningClassTimes[i];
                string eveningName = $"晚修{i + 1}";
                if (DateTime.TryParse(time.Start, out DateTime start) && DateTime.TryParse(time.End, out DateTime end))
                {
                    DateTime startToday = DateTime.Today.Add(start.TimeOfDay);
                    DateTime endToday = DateTime.Today.Add(end.TimeOfDay);
                    if (now >= startToday && now < endToday)
                    {
                        newActive.Add(eveningName);
                    }
                    else if (now >= endToday.AddMinutes(-2) && now <= endToday.AddMinutes(2))
                    {
                        newFlashing.Add(eveningName);
                        newFlashingExists = true;
                    }
                    else if (now > endToday.AddMinutes(2))
                    {
                        newGray.Add(eveningName);
                    }
                }
            }

            bool changed = !_activeEvenings.SequenceEqual(newActive) ||
                           !_flashingEvenings.SequenceEqual(newFlashing) ||
                           !_grayEvenings.SequenceEqual(newGray);

            if (changed)
            {
                _activeEvenings = newActive;
                _flashingEvenings = newFlashing;
                _grayEvenings = newGray;

                _previousFlashingState = newFlashingExists;

                if (_flashingEvenings.Count > 0 && !_debugFlashing)
                {
                    StartFlashing();
                }
                else if (_flashingEvenings.Count == 0 && !_debugFlashing)
                {
                    StopFlashingIfNeeded();
                }
                Invalidate();
            }
        }

        private void StartFlashing()
        {
            if (!flashTimer.Enabled)
            {
                flashStartTime = DateTime.Now;
                flashTimer.Start();
            }
        }

        private void StopFlashingIfNeeded()
        {
            if (_flashingEvenings.Count == 0 && !_debugFlashing)
            {
                flashTimer.Stop();
                Invalidate();
            }
        }

        private void FlashTimer_Tick(object sender, EventArgs e)
        {
            if (_debugFlashing && (DateTime.Now - _debugFlashStartTime).TotalSeconds > FLASH_DURATION)
            {
                _debugFlashing = false;
                StopFlashingIfNeeded();
            }

            double angle = (DateTime.Now - flashStartTime).TotalMilliseconds / 500.0;
            flashStep = (int)((Math.Sin(angle) + 1) * 60);

            _laserOffset += 0.1f;
            if (_laserOffset >= 1f) _laserOffset -= 1f;

            Invalidate();
        }

        public void StartDebugFlashing()
        {
            _debugFlashing = true;
            _debugFlashStartTime = DateTime.Now;
            StartFlashing();
            Invalidate();
        }

        public void StopDebugFlashing()
        {
            _debugFlashing = false;
            StopFlashingIfNeeded();
        }

        // ---------- 滚动 ----------
        private void ScrollTimer_Tick(object sender, EventArgs e)
        {
            if (editMode) return;

            if (appConfig.EnableMarkdown && !rotationMode && !unsubmittedMode)
            {
                float textAreaHeight = gridRects.Count > 0 ? gridRects[0].Height - 70 : 400;
                if (_markdownTotalHeight > textAreaHeight)
                {
                    float mdSpeed = appConfig.ScrollSpeed * 0.05f;
                    float mdEpsilon = 0.01f;

                    if (_markdownScrollPaused)
                    {
                        if ((DateTime.Now - _markdownPauseStart).TotalSeconds >= SCROLL_PAUSE_SECONDS)
                        {
                            if (_markdownScrollOffset >= _markdownTotalHeight - textAreaHeight - mdEpsilon)
                            {
                                _markdownScrollOffset = 0f;
                                _markdownPauseStart = DateTime.Now;
                            }
                            else
                            {
                                _markdownScrollPaused = false;
                                _markdownScrollOffset = mdSpeed;
                            }
                        }
                    }
                    else
                    {
                        if (_markdownScrollOffset <= mdEpsilon)
                        {
                            _markdownScrollPaused = true;
                            _markdownPauseStart = DateTime.Now;
                        }
                        else
                        {
                            _markdownScrollOffset += mdSpeed;
                            if (_markdownScrollOffset >= _markdownTotalHeight - textAreaHeight - mdEpsilon)
                            {
                                _markdownScrollOffset = _markdownTotalHeight - textAreaHeight;
                                _markdownScrollPaused = true;
                                _markdownPauseStart = DateTime.Now;
                            }
                        }
                    }
                    Invalidate();
                }
                return;
            }

            bool needRedraw = false;
            float speed = appConfig.ScrollSpeed * 0.05f;
            float epsilon = 0.01f;

            List<int> indicesToCheck = new List<int>();
            if (rotationMode)
            {
                if (rotationIndex >= 0 && rotationIndex < currentSubjects.Length)
                    indicesToCheck.Add(rotationIndex);
            }
            else if (unsubmittedMode)
            {
                // 未交名单不滚动
            }
            else
            {
                for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
                    indicesToCheck.Add(i);
            }

            foreach (int i in indicesToCheck)
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
                        else
                        {
                            scrollPaused[i] = false;
                            scrollOffsets[i] = speed;
                        }
                    }
                }
                else
                {
                    if (scrollOffsets[i] <= epsilon)
                    {
                        scrollPaused[i] = true;
                        pauseStartTime[i] = DateTime.Now;
                    }
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

        private float MeasureTextHeight(string text, Font font, int maxWidth)
        {
            using (Graphics g = this.CreateGraphics())
            {
                SizeF size = g.MeasureString(text, font, maxWidth);
                return size.Height;
            }
        }

        private Rectangle GetContentArea(int subjectIndex)
        {
            if (rotationMode)
            {
                Rectangle bigRect = new Rectangle(150, 100, 900, 475);
                return new Rectangle(bigRect.Left + 50, bigRect.Top + 150, bigRect.Width - 100, bigRect.Height - 200);
            }
            else
            {
                return new Rectangle(
                    gridRects[subjectIndex].Left + 10,
                    gridRects[subjectIndex].Top + 60,
                    gridRects[subjectIndex].Width - 20,
                    gridRects[subjectIndex].Height - 70);
            }
        }

        // ---------- 编辑模式属性 ----------
        private bool EditMode
        {
            get { return editMode; }
            set
            {
                if (editMode != value)
                {
                    editMode = value;
                    if (editMode)
                    {
                        if (!unsubmittedMode)
                        {
                            CreateTimeComboBoxes();
                        }
                        scrollOffsets.Clear();
                        scrollPaused.Clear();
                        pauseStartTime.Clear();
                    }
                    else
                    {
                        DestroyTimeComboBoxes();
                        if (editingSubjectIndex != -1)
                            FinishInlineEdit();
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

        // ---------- 提交时间下拉框管理 ----------
        private void CreateTimeComboBoxes()
        {
            DestroyTimeComboBoxes();
            for (int i = 0; i < gridRects.Count && i < currentSubjects.Length; i++)
            {
                string subject = currentSubjects[i];
                Rectangle virtualRect = GetDueTimeRect(i);
                Point screenLoc = MapToScreen(virtualRect.Location);
                Size screenSize = new Size(
                    (int)(virtualRect.Width * scaleFactor),
                    (int)(virtualRect.Height * scaleFactor)
                );

                string currentValue = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";

                var comboBox = new ComboBox
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
                List<string> items = new List<string>();
                for (int j = 1; j <= appConfig.EveningClassCount; j++)
                {
                    items.Add($"晚修{j}");
                }
                items.Add("无");
                comboBox.Items.AddRange(items.ToArray());

                int selectedIndex = comboBox.Items.IndexOf(currentValue);
                if (selectedIndex < 0)
                    selectedIndex = appConfig.EveningClassCount;
                comboBox.SelectedIndex = selectedIndex;

                comboBox.DrawItem += TimeComboBox_DrawItem;
                comboBox.SelectedIndexChanged += TimeComboBox_SelectedIndexChanged;

                this.Controls.Add(comboBox);
                timeComboBoxes.Add(comboBox);
            }
        }

        private void DestroyTimeComboBoxes()
        {
            foreach (var combo in timeComboBoxes)
            {
                combo.SelectedIndexChanged -= TimeComboBox_SelectedIndexChanged;
                combo.DrawItem -= TimeComboBox_DrawItem;
                this.Controls.Remove(combo);
                combo.Dispose();
            }
            timeComboBoxes.Clear();
        }

        private void TimeComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            ComboBox combo = sender as ComboBox;
            string text = combo.Items[e.Index].ToString();

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bgColor = isSelected ? Color.FromArgb(80, 80, 80) : Color.FromArgb(40, 40, 40);
            using (SolidBrush bgBrush = new SolidBrush(bgColor))
            {
                e.Graphics.FillRectangle(bgBrush, e.Bounds);
            }

            using (SolidBrush textBrush = new SolidBrush(Color.White))
            {
                e.Graphics.DrawString(text, combo.Font, textBrush, e.Bounds.Left + 5, e.Bounds.Top + 2);
            }

            e.DrawFocusRectangle();
        }

        private void TimeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combo = sender as ComboBox;
            if (combo == null) return;
            int subjectIndex = (int)combo.Tag;
            string subject = currentSubjects[subjectIndex];
            string newValue = combo.SelectedItem?.ToString() ?? "";
            homeworkData.DueTimes[subject] = newValue;
            SaveHomeworkData();
        }

        // ---------- OnLoad 方法 ----------
        private void OnLoad(object sender, EventArgs e)
        {
            ApplyBackgroundEffect(appConfig.BackgroundEffect);

            if (appConfig.UpdatePending == 1)
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string upgradePath = Path.Combine(appData, "HomeworkViewerUpgrader");
                if (Directory.Exists(upgradePath))
                {
                    try
                    {
                        Directory.Delete(upgradePath, true);
                    }
                    catch { }
                }
                appConfig.UpdatePending = 0;
                appConfig.Save();
            }

            var version = Environment.OSVersion.Version;
            if (version.Major == 10 && version.Minor == 0 && version.Build < 22000)
            {
                // 修复 Win10 全屏问题：先设置无边框再最大化
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                fullscreen = true;
                UpdateScale();
            }
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            rotationTimer?.Dispose();
            timeCheckTimer?.Dispose();
            flashTimer?.Dispose();
            scrollTimer?.Dispose();
            font12?.Dispose();
            font20?.Dispose();
            font24?.Dispose();
            font30?.Dispose();
            font22?.Dispose();
            font36?.Dispose();
            hintFont?.Dispose();
            buttonFont?.Dispose();
            fontSmall?.Dispose();
            RED_SEMI?.Dispose();
            ORANGE_SEMI?.Dispose();
            GREEN_SEMI?.Dispose();
            BLUE_SEMI?.Dispose();
            PURPLE_SEMI?.Dispose();
            DARKORANGE_SEMI?.Dispose();
            base.OnFormClosed(e);
        }
    }
}