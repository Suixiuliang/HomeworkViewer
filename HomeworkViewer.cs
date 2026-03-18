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
        private const int BTN_SQUARE_SIZE = 46; // 正方形按钮边长
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
        private Rectangle rotateBtnRect, editBtnRect, historyBtnRect, fullscreenBtnRect, backBtnRect, leftArrowRect, rightArrowRect, settingsBtnRect, unsubmittedBtnRect;

        // 网格矩形
        private List<Rectangle> gridRects = new List<Rectangle>();
        private Rectangle[,] fieldRects;

        // 字体
        private Font font12, font20, font24, font30, font22, font36, hintFont, buttonFont, fontSmall;

        // 颜色
        private Color TEXT_COLOR = Color.Black;
        private readonly Brush RED_SEMI = new SolidBrush(Color.FromArgb(120, 255, 0, 0));
        private readonly Brush ORANGE_SEMI = new SolidBrush(Color.FromArgb(120, 255, 165, 0));
        private readonly Brush GREEN_SEMI = new SolidBrush(Color.FromArgb(120, 0, 255, 0));
        private readonly Brush BLUE_SEMI = new SolidBrush(Color.FromArgb(120, 0, 0, 255));
        private readonly Brush PURPLE_SEMI = new SolidBrush(Color.FromArgb(120, 128, 0, 128));
        private readonly Brush DARKORANGE_SEMI = new SolidBrush(Color.FromArgb(120, 255, 140, 0));

        // 缩放相关
        private float scaleFactor = 1.0f;
        private Point offset = Point.Empty;
        private bool _resizing = false;

        // 图片资源
        private Image buttonImage, historyBtnImage, backBtnImage, leftArrowImage, rightArrowImage;
        private Dictionary<string, Image> buttonIcons = new Dictionary<string, Image>(); // 图标字典

        // 配置
        private AppConfig appConfig;

        // 字体缩放因子
        private float[] fontScales = { 0.8f, 1.0f, 1.2f };
        private string[] currentSubjects;

        // 托盘图标
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;

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

            InitializeTrayIcon();
            InitializeModeComboBox();

            CheckEveningClassStates();
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
            // 编辑模式下不进行任何滚动更新
            if (editMode) return;

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
                // 未交名单模式不滚动
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
                if (contentHeight <= textArea.Height) continue; // 内容无需滚动

                if (!scrollOffsets.ContainsKey(i)) scrollOffsets[i] = 0f;
                if (!scrollPaused.ContainsKey(i)) scrollPaused[i] = false;
                if (!pauseStartTime.ContainsKey(i)) pauseStartTime[i] = DateTime.Now;

                if (speed <= 0) continue;

                if (scrollPaused[i])
                {
                    if ((DateTime.Now - pauseStartTime[i]).TotalSeconds >= SCROLL_PAUSE_SECONDS)
                    {
                        if (scrollOffsets[i] >= contentHeight - textArea.Height - epsilon) // 在底部暂停
                        {
                            // 重置到顶部，并进入顶部暂停
                            scrollOffsets[i] = 0f;
                            pauseStartTime[i] = DateTime.Now;
                            // scrollPaused[i] 保持 true
                        }
                        else // 在顶部暂停
                        {
                            scrollPaused[i] = false;
                            // 给予一个微小偏移量，使其脱离顶部暂停条件
                            scrollOffsets[i] = speed;
                        }
                    }
                }
                else
                {
                    // 未暂停，检查是否已在顶部（可能刚重置或刚解除暂停后被设为了speed）
                    if (scrollOffsets[i] <= epsilon) // 在顶部
                    {
                        // 进入顶部暂停
                        scrollPaused[i] = true;
                        pauseStartTime[i] = DateTime.Now;
                        // 注意：此时scrollOffsets仍为0或接近0
                    }
                    else
                    {
                        // 向下滚动
                        scrollOffsets[i] += speed;
                        if (scrollOffsets[i] >= contentHeight - textArea.Height - epsilon)
                        {
                            scrollOffsets[i] = contentHeight - textArea.Height; // 精确到底
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
                        // 进入编辑模式
                        if (!unsubmittedMode)
                        {
                            CreateTimeComboBoxes();
                        }
                        // 重置滚动状态，使内容显示顶部
                        scrollOffsets.Clear();
                        scrollPaused.Clear();
                        pauseStartTime.Clear();
                    }
                    else
                    {
                        // 退出编辑模式
                        DestroyTimeComboBoxes();
                        if (editingSubjectIndex != -1)
                            FinishInlineEdit();
                        // 重置滚动状态，重新开始
                        scrollOffsets.Clear();
                        scrollPaused.Clear();
                        pauseStartTime.Clear();
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

        // ---------- 初始化托盘图标 ----------
        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("显示", null, (s, e) => ShowWindow());
            trayMenu.Items.Add("退出", null, (s, e) => ExitApplication());

            trayIcon = new NotifyIcon
            {
                ContextMenuStrip = trayMenu,
                Visible = true,
                Text = "作业展板"
            };

            string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Slide", "icon.ico");
            if (File.Exists(iconPath))
            {
                trayIcon.Icon = new Icon(iconPath);
            }
            else
            {
                trayIcon.Icon = SystemIcons.Application;
            }

            trayIcon.MouseDoubleClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                    ShowWindow();
            };
        }

        private void ShowWindow()
        {
            this.Show();
            if (_savedClientSize != Size.Empty)
            {
                this.ClientSize = _savedClientSize;
            }
            this.WindowState = _savedWindowState != FormWindowState.Minimized ? _savedWindowState : FormWindowState.Normal;
            UpdateScale();
            Invalidate();
            this.Activate();
        }

        private void ExitApplication()
        {
            trayIcon.Visible = false;
            trayIcon.Dispose();
            Application.Exit();
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
                trayIcon.ShowBalloonTip(1000, "作业展板", "程序已最小化到托盘，点击图标可恢复。", ToolTipIcon.Info);
            }
        }

        // ---------- 公开方法供设置窗体调用 ----------
        public AppConfig GetConfig() => appConfig;

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
            string fontName = "微软雅黑";

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

        // ---------- Mica 效果 ----------
        private void OnLoad(object sender, EventArgs e)
        {
            EnableMica();
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

        private void EnableAcrylicFallback()
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
            float rectWidth = (areaWidth - (cols + 1) * GRID_PADDING) / cols;
            float rectHeight = (areaHeight - (rows + 1) * GRID_PADDING) / rows;

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

            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _unsubmittedPressed = false;

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
            }
            else
            {
                if (settingsBtnRect.Contains(x, y))
                    _settingsPressed = true;
                else if (rotateBtnRect.Contains(x, y))
                    _rotatePressed = true;
                else if (unsubmittedBtnRect.Contains(x, y))
                    _unsubmittedPressed = true;
                else if (editBtnRect.Contains(x, y))
                    _editPressed = true;
                else if (historyBtnRect.Contains(x, y))
                    _historyPressed = true;
                else if (fullscreenBtnRect.Contains(x, y))
                    _fullscreenPressed = true;
            }
            Invalidate();
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _unsubmittedPressed = false;
            Invalidate();
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = _unsubmittedPressed = false;
            Invalidate();
        }

        private void OnResize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                _savedClientSize = this.ClientSize;
                _savedWindowState = this.WindowState;
                // 移除 Hide() 和托盘提示，让窗口正常最小化
                return; // 直接返回，不执行后续缩放逻辑
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
            Invalidate();
        }

        private void OnActivated(object sender, EventArgs e)
        {
            UpdateScale();
            Invalidate();
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x0112;
            const int SC_MAXIMIZE = 0xF030;
            const int WM_SIZE = 0x0005;
            const int SIZE_RESTORED = 0;

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
                Invalidate();
                if (!_micaEnabled)
                {
                    EnableMica();
                }
            }

            base.WndProc(ref m);
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
                    // 全屏前先退出编辑模式
                    if (EditMode)
                        EditMode = false;
                    ToggleFullscreen();
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
                    // 全屏前先退出编辑模式
                    if (EditMode)
                        EditMode = false;
                    ToggleFullscreen();
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
                if (settingsBtnRect.Contains(x, y))
                {
                    using (var settingsForm = new SettingsForm(this))
                    {
                        settingsForm.ShowDialog();
                    }
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
                    // 全屏前先退出编辑模式
                    if (EditMode)
                        EditMode = false;
                    ToggleFullscreen();
                }
                else if (EditMode && !rotationMode && !unsubmittedMode)
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

            var textBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Location = screenLoc,
                Size = screenSize,
                Text = currentText,
                Font = font30,
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
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Slide");
            // 原有图片
            buttonImage = LoadImage(Path.Combine(imagePath, "按钮.png"), new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));
            historyBtnImage = LoadImage(Path.Combine(imagePath, "更多.png"), new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));

            Image originalBack = LoadImage(Path.Combine(imagePath, "返回.png"), originalSize: true);
            if (originalBack != null)
            {
                int targetHeight = BTN_SQUARE_SIZE;
                int targetWidth = (int)((float)originalBack.Width / originalBack.Height * targetHeight);
                backBtnImage = new Bitmap(originalBack, new Size(targetWidth, targetHeight));
            }

            leftArrowImage = LoadImage(Path.Combine(imagePath, "箭头图片.png"), new Size(31, 50));
            if (leftArrowImage != null)
            {
                rightArrowImage = (Image)new Bitmap(leftArrowImage);
                rightArrowImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            }

            // 加载所有按钮图标（高质量缩放）
            string[] iconNames = { "编辑", "返回", "记录", "轮换", "全屏", "设置", "缩小", "完成", "未交人员" };
            int iconSize = (int)(BTN_SQUARE_SIZE * 0.6); // 图标占按钮高度的60%，避免重叠
            foreach (string name in iconNames)
            {
                string filePath = Path.Combine(imagePath, name + ".png");
                if (File.Exists(filePath))
                {
                    using (Image img = Image.FromFile(filePath))
                    {
                        // 创建目标位图，保留透明度
                        Bitmap targetBitmap = new Bitmap(iconSize, iconSize, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                        using (Graphics g = Graphics.FromImage(targetBitmap))
                        {
                            // 设置高质量插值模式
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.SmoothingMode = SmoothingMode.HighQuality;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            g.CompositingQuality = CompositingQuality.HighQuality;
                            // 绘制缩放后的图像
                            g.DrawImage(img, 0, 0, iconSize, iconSize);
                        }
                        buttonIcons[name] = targetBitmap;
                    }
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
                // 左上角返回按钮
                int backY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;
                backBtnRect = new Rectangle(35, backY, BTN_SQUARE_SIZE, BTN_SQUARE_SIZE);

                DrawTransparentButton(g, backBtnRect, "返回", buttonIcons.ContainsKey("返回") ? buttonIcons["返回"] : null, ref backBtnRect, backBtnRect.Location, _backPressed);

                // 右上角三个按钮
                int btnY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;
                int totalButtonsWidth = 3 * BTN_SQUARE_SIZE + 2 * 10;
                int startX = barRect.Right - totalButtonsWidth - 20;

                Point editPos = new Point(startX, btnY);
                Point historyPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point fullscreenPos = new Point(historyPos.X + BTN_SQUARE_SIZE + 10, btnY);

                string editButtonText = unsubmittedMode ? (editMode ? "完成" : "登记") : (editMode ? "完成" : "编辑");
                string editIconKey = editMode ? "完成" : (unsubmittedMode ? "编辑" : "编辑");
                DrawTransparentButton(g, editBtnRect, editButtonText, buttonIcons.ContainsKey(editIconKey) ? buttonIcons[editIconKey] : null, ref editBtnRect, editPos, _editPressed);

                string historyIconKey = historyMode ? "返回" : "记录";
                DrawTransparentButton(g, historyBtnRect, historyMode ? "返回" : "记录", buttonIcons.ContainsKey(historyIconKey) ? buttonIcons[historyIconKey] : null, ref historyBtnRect, historyPos, _historyPressed);

                string fullscreenIconKey = fullscreen ? "缩小" : "全屏";
                DrawTransparentButton(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreenIconKey) ? buttonIcons[fullscreenIconKey] : null, ref fullscreenBtnRect, fullscreenPos, _fullscreenPressed);

                // 绘制左右箭头（如果分页）
                if (unsubmittedMode)
                {
                    int totalPages = (currentSubjects.Length + UNSUBMITTED_ROWS_PER_PAGE - 1) / UNSUBMITTED_ROWS_PER_PAGE;
                    if (totalPages > 1)
                    {
                        leftArrowRect = new Rectangle(50, VIRTUAL_SIZE.Height / 2 - 25, 31, 50);
                        rightArrowRect = new Rectangle(VIRTUAL_SIZE.Width - 50 - 31, VIRTUAL_SIZE.Height / 2 - 25, 31, 50);
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
                    leftArrowRect = new Rectangle(50, VIRTUAL_SIZE.Height / 2 - 25, 31, 50);
                    rightArrowRect = new Rectangle(VIRTUAL_SIZE.Width - 50 - 31, VIRTUAL_SIZE.Height / 2 - 25, 31, 50);
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
                int btnY = barRect.Top + (barRect.Height - BTN_SQUARE_SIZE) / 2;
                int totalButtonsWidth = 6 * BTN_SQUARE_SIZE + 5 * 10;
                int startX = barRect.Right - totalButtonsWidth - 20;

                Point settingsPos = new Point(startX, btnY);
                Point rotatePos = new Point(settingsPos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point unsubmittedPos = new Point(rotatePos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point editPos = new Point(unsubmittedPos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point historyPos = new Point(editPos.X + BTN_SQUARE_SIZE + 10, btnY);
                Point fullscreenPos = new Point(historyPos.X + BTN_SQUARE_SIZE + 10, btnY);

                DrawTransparentButton(g, settingsBtnRect, "设置", buttonIcons.ContainsKey("设置") ? buttonIcons["设置"] : null, ref settingsBtnRect, settingsPos, _settingsPressed);
                DrawTransparentButton(g, rotateBtnRect, "轮换", buttonIcons.ContainsKey("轮换") ? buttonIcons["轮换"] : null, ref rotateBtnRect, rotatePos, _rotatePressed);
                DrawTransparentButton(g, unsubmittedBtnRect, "未交", buttonIcons.ContainsKey("未交人员") ? buttonIcons["未交人员"] : null, ref unsubmittedBtnRect, unsubmittedPos, _unsubmittedPressed);
                DrawTransparentButton(g, editBtnRect, editMode ? "完成" : "编辑", buttonIcons.ContainsKey(editMode ? "完成" : "编辑") ? buttonIcons[editMode ? "完成" : "编辑"] : null, ref editBtnRect, editPos, _editPressed);
                DrawTransparentButton(g, historyBtnRect, historyMode ? "返回" : "记录", buttonIcons.ContainsKey(historyMode ? "返回" : "记录") ? buttonIcons[historyMode ? "返回" : "记录"] : null, ref historyBtnRect, historyPos, _historyPressed);
                DrawTransparentButton(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", buttonIcons.ContainsKey(fullscreen ? "缩小" : "全屏") ? buttonIcons[fullscreen ? "缩小" : "全屏"] : null, ref fullscreenBtnRect, fullscreenPos, _fullscreenPressed);
            }
        }

        private void DrawTransparentButton(Graphics g, Rectangle rect, string text, Image icon, ref Rectangle targetRect, Point pos, bool pressed)
        {
            targetRect = new Rectangle(pos, new Size(BTN_SQUARE_SIZE, BTN_SQUARE_SIZE));

            // 可调参数（可根据需要修改）
            const int iconTopMargin = 3;      // 图标上边距（像素）
            const float textBottomMargin = 1; // 文字下边距（像素，原为3，现增加2达到下移效果）

            // 绘制图标
            if (icon != null)
            {
                int iconX = targetRect.Left + (targetRect.Width - icon.Width) / 2;
                int iconY = targetRect.Top + iconTopMargin;
                g.DrawImage(icon, iconX, iconY, icon.Width, icon.Height);
            }

            // 绘制文字
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

                // 阴影
                int shadowOffset = 3;
                Rectangle shadowRect = new Rectangle(rect.X + shadowOffset, rect.Y + shadowOffset, rect.Width, rect.Height);
                using (GraphicsPath shadowPath = CreateRoundedRectPath(shadowRect, ROUND_RADIUS))
                {
                    using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(30, 0, 0, 0)))
                    {
                        g.FillPath(shadowBrush, shadowPath);
                    }
                }

                // 卡片主体背景（毛玻璃渐变）
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
                    Color borderColor = Color.FromArgb(100, 128, 128, 128); // 默认边框颜色
                    float borderWidth = 1; // 默认边框宽度

                    if (_debugFlashing || highlightFlash)
                    {
                        int alpha = flashStep;
                        using (SolidBrush redBrush = new SolidBrush(Color.FromArgb(alpha, 255, 0, 0)))
                        {
                            g.FillPath(redBrush, path);
                        }
                        DrawLaserBorder(g, path, rect);
                        borderColor = Color.FromArgb(200, 255, 0, 0);
                        // 闪烁时边框宽度保持1，激光已单独处理
                    }
                    else if (highlightActive)
                    {
                        // 晚修进行时：不填充任何额外背景，仅边框加粗为2，颜色使用系统高亮色
                        borderColor = SystemColors.Highlight;
                        borderWidth = 2;
                    }
                    // highlightGray 等其他情况不处理，保持默认

                    // 绘制卡片边框
                    using (Pen borderPen = new Pen(borderColor, borderWidth))
                    {
                        g.DrawPath(borderPen, path);
                    }
                }

                // 绘制科目名称、横线、提交时间、作业内容等（以下代码保持不变）
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
                    // 晚修进行时：不填充额外背景，仅边框加粗为2，颜色使用系统高亮色
                    borderColor = SystemColors.Highlight;
                    borderWidth = 2;
                }

                using (Pen borderPen = new Pen(borderColor, borderWidth))
                {
                    g.DrawPath(borderPen, path);
                }
            }

            // 以下为卡片内部其他内容绘制（保持不变）
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
            // 确保字体有效
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

                // 定义间距常量
                const int subjectToCountSpacing = 2;   // 科目与人数统计间距
                const int countToListSpacing = 5;      // 人数统计与名单区域间距
                const int listVerticalPadding = -300;      // 名单区域上下内边距

                for (int i = startIndex; i < endIndex; i++)
                {
                    string subject = currentSubjects[i];
                    int rowY = baseY + (i - startIndex) * rowHeight;

                    // 绘制科目名称（左对齐）
                    float subjectY = rowY + 5;
                    g.DrawString(subject, font22, new SolidBrush(TEXT_COLOR), new PointF(bigRect.Left + 50, subjectY));

                    // 计算人数
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

                    // 绘制人数统计（红色）
                    string countText = peopleCount.ToString();
                    string suffixText = "人未交";
                    float countWidth = g.MeasureString(countText, font30).Width;
                    float statsX = bigRect.Left + 50;
                    float statsY = subjectY + font22.Height + subjectToCountSpacing;
                    g.DrawString(countText, font30, Brushes.Red, statsX, statsY);
                    g.DrawString(suffixText, fontSmall, Brushes.Red, statsX + countWidth, statsY + font30.Height - fontSmall.Height - 2);

                    // 计算名单区域
                    float usedHeight = (subjectY - rowY) + font22.Height + subjectToCountSpacing + font30.Height; // 从 rowY 到人数统计底部的高度
                    float availableForList = rowHeight - usedHeight - countToListSpacing; // 减去人数统计与名单之间的间距
                    if (availableForList < listVerticalPadding * 2) availableForList = listVerticalPadding * 2; // 保证至少有内边距空间

                    float listAreaHeight = availableForList - listVerticalPadding * 2; // 实际绘制文本的高度
                    float listAreaTop = rowY + usedHeight + countToListSpacing + listVerticalPadding - 40; // 名单区域顶部

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
                        // 有名单文字，需要分行并垂直居中
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

                        // 计算垂直偏移，使文本在 textRect 内居中
                        float verticalOffset = (textRect.Height - totalTextHeight) / 2;
                        if (verticalOffset < 0) verticalOffset = 0; // 超出时不滚动，显示顶部

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
                        // 无名单文字，显示提示，手动居中
                        string hint = editMode ? "点此编辑名单" : "无未交人员";
                        SizeF hintSize = g.MeasureString(hint, hintFont);
                        float hintX = textRect.Left + (textRect.Width - hintSize.Width) / 2;
                        float hintY = textRect.Top + (textRect.Height - hintSize.Height) / 2;
                        g.DrawString(hint, hintFont, RED_SEMI, hintX, hintY);
                    }

                    // 绘制行分隔线（最后一行不画）
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

        private void DrawTextInArea(Graphics g, string text, Rectangle area, Font font, bool center = false, float scrollOffset = 0f)
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
                        if (line != "") lines.Add(line);
                        line = word;
                    }
                }
                if (line != "") lines.Add(line);
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
            trayIcon?.Dispose();
            base.OnFormClosed(e);
        }
    }
}