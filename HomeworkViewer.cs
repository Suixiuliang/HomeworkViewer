#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace HomeworkViewer
{
    public class HomeworkViewer : Form
    {
        // 常量
        private readonly Size VIRTUAL_SIZE = new Size(1200, 675);
        private readonly Size BTN_SIZE = new Size(60, 20);
        private readonly Point FULLSCREEN_BTN_POS = new Point(1200 - 60 - 10, 13);
        private readonly Point HISTORY_BTN_POS = new Point(1200 - 60 - 10 - 60 - 10, 13);
        private readonly Point EDIT_BTN_POS = new Point(1200 - 60 - 10 - 60 - 10 - 60 - 10, 13);
        private readonly Point ROTATE_BTN_POS = new Point(1200 - 60 - 10 - 60 - 10 - 60 - 10 - 60 - 10, 13);
        private Point SETTINGS_BTN_POS;
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
        private int rotationIndex = 0;
        private Timer rotationTimer;
        private const int rotationInterval = 10000;

        // 内联编辑（用于科目内容）
        private Control inlineEditControl;
        private int editingSubjectIndex = -1;
        private enum EditFieldType { None, Subject, DueTime }
        private EditFieldType currentEditType = EditFieldType.None;

        // 数据
        private HomeworkData homeworkData = new HomeworkData();
        private DateTime currentDate = DateTime.Now;
        private DateTime? historyDate = null;

        // 按钮矩形
        private Rectangle rotateBtnRect, editBtnRect, historyBtnRect, fullscreenBtnRect, backBtnRect, leftArrowRect, rightArrowRect, settingsBtnRect;

        // 网格矩形
        private List<Rectangle> gridRects = new List<Rectangle>();
        private Rectangle[,] fieldRects;

        // 字体
        private Font font12, font20, font24, font30, font22, font36, hintFont, buttonFont, fontSmall;

        // 颜色
        private Color TEXT_COLOR = Color.Black;
        // 提交时间颜色画刷（固定透明度120）
        private readonly Brush RED_SEMI = new SolidBrush(Color.FromArgb(120, 255, 0, 0));           // 晚修一红色
        private readonly Brush ORANGE_SEMI = new SolidBrush(Color.FromArgb(120, 255, 165, 0));     // 晚修二橙色
        private readonly Brush GREEN_SEMI = new SolidBrush(Color.FromArgb(120, 0, 255, 0));         // 晚修三鲜艳绿色

        // 缩放相关
        private float scaleFactor = 1.0f;
        private Point offset = Point.Empty;
        private bool _resizing = false;

        // 图片资源
        private Image buttonImage, historyBtnImage, backBtnImage, leftArrowImage, rightArrowImage;

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

        // 模式切换下拉框
        private ComboBox modeComboBox;

        // 编辑模式下的提交时间下拉框列表
        private List<ComboBox> timeComboBoxes = new List<ComboBox>();

        public HomeworkViewer()
        {
            SETTINGS_BTN_POS = new Point(ROTATE_BTN_POS.X - 60 - 10, 13);

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
                        CreateTimeComboBoxes();
                    }
                    else
                    {
                        DestroyTimeComboBoxes();
                        if (editingSubjectIndex != -1)
                            FinishInlineEdit();
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
                if (string.IsNullOrEmpty(currentValue))
                    currentValue = "晚修三";

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
                // 移除"明天早读"
                comboBox.Items.AddRange(new object[] { "晚修一", "晚修二", "晚修三", "无" });
                comboBox.SelectedItem = currentValue;
                if (comboBox.SelectedItem == null)
                    comboBox.SelectedIndex = 2;
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

        // Mica API
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
            if (EditMode)
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

            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = false;

            if (rotationMode)
            {
                if (backBtnRect.Contains(x, y))
                    _backPressed = true;
            }
            else
            {
                if (settingsBtnRect.Contains(x, y))
                    _settingsPressed = true;
                else if (rotateBtnRect.Contains(x, y))
                    _rotatePressed = true;
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
            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = false;
            Invalidate();
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            _settingsPressed = _rotatePressed = _editPressed = _historyPressed = _fullscreenPressed = _backPressed = false;
            Invalidate();
        }

        private void OnResize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                _savedClientSize = this.ClientSize;
                _savedWindowState = this.WindowState;
                this.Hide();
                trayIcon.ShowBalloonTip(1000, "作业展板", "程序已最小化到托盘，点击图标可恢复。", ToolTipIcon.Info);
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
            if (EditMode)
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
            string displayTime = string.IsNullOrEmpty(dueTime) ? "晚修三" : dueTime;

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
            if (EditMode)
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
                    ToggleFullscreen();
                }
                else if (EditMode && !rotationMode)
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

        // ---------- 内联编辑（科目内容） ----------
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

            string currentText = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
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
            homeworkData.Subjects[subject] = newText;
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
            buttonImage = LoadImage(Path.Combine(imagePath, "按钮.png"), BTN_SIZE);
            historyBtnImage = LoadImage(Path.Combine(imagePath, "更多.png"), BTN_SIZE);

            Image originalBack = LoadImage(Path.Combine(imagePath, "返回.png"), originalSize: true);
            if (originalBack != null)
            {
                int targetHeight = 37;
                int targetWidth = (int)((float)originalBack.Width / originalBack.Height * targetHeight);
                backBtnImage = new Bitmap(originalBack, new Size(targetWidth, targetHeight));
            }

            leftArrowImage = LoadImage(Path.Combine(imagePath, "箭头图片.png"), new Size(31, 50));
            if (leftArrowImage != null)
            {
                rightArrowImage = (Image)new Bitmap(leftArrowImage);
                rightArrowImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
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
            switch (value)
            {
                case "晚修一": return RED_SEMI;
                case "晚修二": return ORANGE_SEMI;
                case "晚修三": return GREEN_SEMI;
                default: return new SolidBrush(TEXT_COLOR);
            }
        }

        // ---------- 绘制 ----------
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            g.TranslateTransform(offset.X, offset.Y);
            g.ScaleTransform(scaleFactor, scaleFactor);

            // 背景
            if (!_micaEnabled)
            {
                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(_backgroundAlpha, 32, 32, 32)))
                {
                    g.FillRectangle(bgBrush, new Rectangle(0, 0, VIRTUAL_SIZE.Width, VIRTUAL_SIZE.Height));
                }
            }

            // 计算动态 Bar 矩形
            Rectangle barRect = new Rectangle(0, BAR_Y, VIRTUAL_SIZE.Width, BAR_HEIGHT);
            if (gridRects.Count >= 3)
            {
                int barLeft = gridRects[0].Left;
                int barRight = gridRects[2].Right;
                barRect = new Rectangle(barLeft, BAR_Y, barRight - barLeft, BAR_HEIGHT);
            }

            // 无论是否轮播，都绘制顶部 Bar
            float opacityFactor = appConfig.CardOpacity / 100f;
            int barAlpha = (int)(255 * opacityFactor);
            Color barColor = ParseColor(appConfig.BarColor, Color.Yellow);
            using (SolidBrush barBrush = new SolidBrush(Color.FromArgb(barAlpha, barColor.R, barColor.G, barColor.B)))
            using (GraphicsPath barPath = CreateRoundedRectPath(barRect, ROUND_RADIUS, bottomOnly: true))
            {
                g.FillPath(barBrush, barPath);
                // 添加边框（半透明白色）
                using (Pen borderPen = new Pen(Color.FromArgb(120, 255, 255, 255), 1))
                {
                    g.DrawPath(borderPen, barPath);
                }
            }

            DrawDateInfo(g, barRect);
            if (rotationMode)
                DrawRotationView(g);
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
                if (rotationMode)
                {
                    // 轮播模式：日期在 Bar 内垂直居中，并水平居中
                    int totalHeight = font20.Height + font12.Height;
                    int startY = barRect.Top + (barRect.Height - totalHeight) / 2;

                    SizeF weekdaySize = g.MeasureString(weekdayText, font20);
                    g.DrawString(weekdayText, font20, whiteBrush, new PointF((VIRTUAL_SIZE.Width - weekdaySize.Width) / 2, startY));

                    SizeF dateSize = g.MeasureString(dateText, font12);
                    g.DrawString(dateText, font12, whiteBrush, new PointF((VIRTUAL_SIZE.Width - dateSize.Width) / 2, startY + font20.Height));
                }
                else
                {
                    // 普通模式：日期在 Bar 内垂直居中，左对齐
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
            if (rotationMode)
            {
                // 返回按钮 X 坐标改为 35
                int backY = 13;
                backBtnRect = new Rectangle(35, backY, BTN_SIZE.Width, BTN_SIZE.Height);

                Color backColor = _backPressed ? Color.FromArgb(200, 40, 40, 40) : Color.FromArgb(180, 64, 64, 64);
                using (GraphicsPath path = CreateRoundedRectPath(backBtnRect, ROUND_RADIUS))
                {
                    using (SolidBrush btnBrush = new SolidBrush(backColor))
                    {
                        g.FillPath(btnBrush, path);
                    }
                    using (Pen borderPen = new Pen(Color.FromArgb(100, 255, 255, 255)))
                    {
                        g.DrawPath(borderPen, path);
                    }
                }
                g.DrawString("返回", buttonFont, Brushes.White, backBtnRect, CenterStringFormat());

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
            else
            {
                int btnY = barRect.Top + (barRect.Height - BTN_SIZE.Height) / 2;
                int totalButtonsWidth = 5 * BTN_SIZE.Width + 4 * 10; // 340
                int startX = barRect.Right - totalButtonsWidth - 20; // 右边距20

                Point settingsPos = new Point(startX, btnY);
                Point rotatePos = new Point(settingsPos.X + BTN_SIZE.Width + 10, btnY);
                Point editPos = new Point(rotatePos.X + BTN_SIZE.Width + 10, btnY);
                Point historyPos = new Point(editPos.X + BTN_SIZE.Width + 10, btnY);
                Point fullscreenPos = new Point(historyPos.X + BTN_SIZE.Width + 10, btnY);

                DrawTransparentButton(g, settingsBtnRect, "设置", ref settingsBtnRect, settingsPos, _settingsPressed);
                DrawTransparentButton(g, rotateBtnRect, "轮换", ref rotateBtnRect, rotatePos, _rotatePressed);
                DrawTransparentButton(g, editBtnRect, editMode ? "完成" : "编辑", ref editBtnRect, editPos, _editPressed);
                DrawTransparentButton(g, historyBtnRect, historyMode ? "返回" : "记录", ref historyBtnRect, historyPos, _historyPressed);
                DrawTransparentButton(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏", ref fullscreenBtnRect, fullscreenPos, _fullscreenPressed);
            }
        }

        private void DrawTransparentButton(Graphics g, Rectangle rect, string text, ref Rectangle targetRect, Point pos, bool pressed)
        {
            targetRect = new Rectangle(pos, BTN_SIZE);
            Color btnColor = pressed ? Color.FromArgb(200, 40, 40, 40) : Color.FromArgb(180, 64, 64, 64);
            using (GraphicsPath path = CreateRoundedRectPath(targetRect, ROUND_RADIUS))
            {
                using (SolidBrush btnBrush = new SolidBrush(btnColor))
                {
                    g.FillPath(btnBrush, path);
                }
                using (Pen borderPen = new Pen(Color.FromArgb(100, 255, 255, 255)))
                {
                    g.DrawPath(borderPen, path);
                }
            }
            g.DrawString(text, buttonFont, Brushes.White, targetRect, CenterStringFormat());
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
                    using (Pen borderPen = new Pen(Color.FromArgb(100, 128, 128, 128), 1))
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

                    // 非编辑模式才绘制时间文本
                    if (!EditMode)
                    {
                        string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
                        string prefix = "提交时间：";
                        string displayTime;
                        if (string.IsNullOrEmpty(dueTime))
                            displayTime = "晚修三";
                        else
                            displayTime = dueTime;

                        float prefixWidth = g.MeasureString(prefix, fontSmall).Width;
                        float timeWidth = g.MeasureString(displayTime, font22).Width;
                        float totalWidth = prefixWidth + timeWidth;
                        int rightX = rect.Right - 10;
                        float startX = rightX - totalWidth;
                        int timeY = lineY - font22.Height;
                        int prefixY = lineY - fontSmall.Height;

                        g.DrawString(prefix, fontSmall, new SolidBrush(TEXT_COLOR), startX, prefixY);
                        // 根据值选择画刷
                        Brush timeBrush = GetDueTimeBrush(displayTime);
                        g.DrawString(displayTime, font22, timeBrush, startX + prefixWidth, timeY);
                        if (timeBrush != RED_SEMI && timeBrush != ORANGE_SEMI && timeBrush != GREEN_SEMI)
                            timeBrush.Dispose(); // 只有临时创建的画刷需要释放
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
                        DrawTextInArea(g, homeworkData.Subjects[subject], textArea, font30);
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
                using (Pen borderPen = new Pen(Color.FromArgb(100, 128, 128, 128), 1))
                {
                    g.DrawPath(borderPen, path);
                }
            }

            string subject = currentSubjects[rotationIndex];
            g.DrawString(subject, font36, Brushes.White, new PointF(600 - (int)g.MeasureString(subject, font36).Width / 2, bigRect.Top + 30));

            int lineY = bigRect.Top + 120;
            g.DrawLine(new Pen(Color.Gray, 2), bigRect.Left + 50, lineY, bigRect.Right - 50, lineY);

            string dueTime = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
            string displayTime;
            if (string.IsNullOrEmpty(dueTime))
                displayTime = "晚修三";
            else
                displayTime = dueTime;

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
            if (timeBrush != RED_SEMI && timeBrush != ORANGE_SEMI && timeBrush != GREEN_SEMI)
                timeBrush.Dispose();

            Rectangle textArea = new Rectangle(bigRect.Left + 50, bigRect.Top + 150, bigRect.Width - 100, bigRect.Height - 200);
            if (homeworkData.Subjects.ContainsKey(subject) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[subject]))
            {
                DrawTextInArea(g, homeworkData.Subjects[subject], textArea, font30, center: true);
            }
            else
            {
                g.DrawString("今天暂时没有此项作业", font30, RED_SEMI, textArea, CenterStringFormat());
            }
        }

        private void DrawTextInArea(Graphics g, string text, Rectangle area, Font font, bool center = false)
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
            int maxLines = (int)(area.Height / lineHeight);
            for (int i = 0; i < Math.Min(lines.Count, maxLines); i++)
            {
                float y = area.Top + i * lineHeight;
                if (center)
                {
                    SizeF sz = g.MeasureString(lines[i], font);
                    g.DrawString(lines[i], font, new SolidBrush(TEXT_COLOR), new PointF(area.Left + (area.Width - sz.Width) / 2, y));
                }
                else
                {
                    g.DrawString(lines[i], font, new SolidBrush(TEXT_COLOR), new PointF(area.Left + 5, y));
                }
            }
            if (lines.Count > maxLines)
            {
                g.DrawString("...", font, new SolidBrush(TEXT_COLOR), new PointF(area.Left + 5, area.Top + maxLines * lineHeight));
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            rotationTimer?.Dispose();
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
            trayIcon?.Dispose();
            base.OnFormClosed(e);
        }
    }
}