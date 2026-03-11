#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
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
        private readonly Rectangle GRID_AREA = new Rectangle(0, 47, 1200, 631);
        private const int GRID_COLS = 3, GRID_ROWS = 2, GRID_PADDING = 20, GRID_BORDER = 1;

        // 状态
        private bool fullscreen = false;
        private bool historyMode = false;
        private bool editMode = false;
        private bool rotationMode = false;
        private int rotationIndex = 0;
        private Timer rotationTimer;
        private const int rotationInterval = 10000;

        // 内联编辑
        private Control inlineEditControl;
        private int editingSubjectIndex = -1;
        private enum EditFieldType { None, Subject, DueTime }
        private EditFieldType currentEditType = EditFieldType.None;

        // 数据
        private HomeworkData homeworkData = new HomeworkData();
        private DateTime currentDate = DateTime.Now;
        private DateTime? historyDate = null;

        // 按钮矩形
        private Rectangle rotateBtnRect, editBtnRect, historyBtnRect, fullscreenBtnRect, backBtnRect, leftArrowRect, rightArrowRect;

        // 网格矩形
        private List<Rectangle> gridRects = new List<Rectangle>();
        private Rectangle[,] fieldRects;

        // 字体
        private Font font12, font20, font24, font30, font22, font36, hintFont, buttonFont, fontSmall;

        // 颜色
        private readonly Color WHITE = Color.White;
        private readonly Color GRAY = Color.Gray;
        private readonly Color BLACK = Color.Black;
        private readonly Color RED = Color.Red;

        // 缩放相关
        private float scaleFactor = 1.0f;
        private Point offset = Point.Empty;
        private bool _resizing = false;

        // 图片资源
        private Image bgImage, dialogBgImage, buttonImage, historyBtnImage, backBtnImage, leftArrowImage, rightArrowImage;

        // 模式切换
        private ComboBox modeComboBox;
        private Dictionary<string, string[]> subjectModes;
        private string[] currentSubjects;
        private string currentModeKey;

        // 配置
        private AppConfig appConfig;

        // 记录全屏前的下拉框可见状态
        private bool _modeComboBoxVisibleBeforeFullscreen;

        public HomeworkViewer()
        {
            Text = "高一三班作业";
            ClientSize = VIRTUAL_SIZE;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;
            DoubleBuffered = true;
            KeyPreview = true; // 允许窗体优先处理按键

            // 初始化字体
            font12 = new Font("微软雅黑", 10);
            font20 = new Font("微软雅黑", 15);
            font24 = new Font("微软雅黑", 14);
            font30 = new Font("微软雅黑", 20);
            font22 = new Font("微软雅黑", 21);
            font36 = new Font("微软雅黑", 35);
            hintFont = new Font("微软雅黑", 15);
            buttonFont = new Font("微软雅黑", 10);
            fontSmall = new Font("微软雅黑", 8);

            // 定义科目模式
            subjectModes = new Dictionary<string, string[]>
            {
                {"大理", new[]{"语文","数学","英语","物理","化学","生物"}},
                {"中理", new[]{"语文","数学","英语","物理","化学","地理"}},
                {"小理", new[]{"语文","数学","英语","物理","化学","政治"}},
                {"大文", new[]{"语文","数学","英语","政治","历史","地理"}},
                {"全科", new[]{"语文","数学","英语","物理","化学","生物","政治","历史","地理"}}
            };

            // 加载配置
            appConfig = AppConfig.Load();
            currentModeKey = appConfig.LastMode;
            if (!subjectModes.ContainsKey(currentModeKey))
                currentModeKey = "大理";

            currentSubjects = subjectModes[currentModeKey];

            // 创建模式选择下拉框
            modeComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 10),
                Width = 100,
                Height = 25,
                Visible = false // 默认隐藏
            };
            modeComboBox.Items.AddRange(new object[] { "大理", "中理", "小理", "大文", "全科" });
            modeComboBox.SelectedItem = currentModeKey;
            modeComboBox.SelectedIndexChanged += ModeComboBox_SelectedIndexChanged;
            this.Controls.Add(modeComboBox);

            // 计算网格
            CalculateGridRects();

            // 加载今日作业
            LoadHomeworkData(currentDate);

            // 加载图片
            LoadImages();

            // 轮播定时器
            rotationTimer = new Timer { Interval = rotationInterval };
            rotationTimer.Tick += (s, e) => { if (rotationMode) RotateNext(); };

            // 事件
            this.MouseClick += OnMouseClick;
            this.Resize += OnResize;
            this.KeyDown += OnKeyDown; // 处理 F8

            // 初始定位下拉框
            UpdateComboBoxPosition();
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F8)
            {
                // 切换下拉框可见性
                modeComboBox.Visible = !modeComboBox.Visible;
                if (modeComboBox.Visible)
                {
                    // 如果显示，确保位置正确并可选地给予焦点
                    UpdateComboBoxPosition();
                    modeComboBox.BringToFront();
                    modeComboBox.Focus(); // 让下拉框获得焦点，方便直接操作
                }
                e.Handled = true;
            }
        }

        private void ModeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string newMode = modeComboBox.SelectedItem.ToString();
            if (newMode == currentModeKey) return;

            currentModeKey = newMode;
            currentSubjects = subjectModes[currentModeKey];

            // 保存配置
            appConfig.LastMode = currentModeKey;
            appConfig.Save();

            CalculateGridRects();
            Invalidate();
        }

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
        }

        private void SaveHomeworkData()
        {
            DateTime saveDate = historyMode && historyDate.HasValue ? historyDate.Value : currentDate;
            homeworkData.Save(saveDate);
        }

        private void ToggleFullscreen()
        {
            if (!fullscreen)
            {
                // 进入全屏前保存下拉框可见状态
                _modeComboBoxVisibleBeforeFullscreen = modeComboBox.Visible;
            }

            fullscreen = !fullscreen;
            if (fullscreen)
            {
                FormBorderStyle = FormBorderStyle.None;
                WindowState = FormWindowState.Maximized;
                // 全屏时强制隐藏下拉框
                modeComboBox.Visible = false;
            }
            else
            {
                FormBorderStyle = FormBorderStyle.Sizable;
                WindowState = FormWindowState.Normal;
                ClientSize = VIRTUAL_SIZE;
                // 退出全屏后恢复之前的可见状态
                modeComboBox.Visible = _modeComboBoxVisibleBeforeFullscreen;
            }
            UpdateScale();
            UpdateComboBoxPosition();
            Invalidate();
        }

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

        private void UpdateComboBoxPosition()
        {
            int comboX = ROTATE_BTN_POS.X - modeComboBox.Width - 10;
            Point virtualPos = new Point(comboX, ROTATE_BTN_POS.Y);
            Point screenPos = MapToScreen(virtualPos);
            modeComboBox.Location = screenPos;
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

        private void OnResize(object sender, EventArgs e)
        {
            if (_resizing || fullscreen || WindowState == FormWindowState.Maximized) return;
            _resizing = true;
            float targetRatio = (float)VIRTUAL_SIZE.Width / VIRTUAL_SIZE.Height;
            int newWidth = this.ClientSize.Width;
            int newHeight = (int)(newWidth / targetRatio);
            this.ClientSize = new Size(newWidth, newHeight);
            _resizing = false;

            UpdateScale();
            UpdateComboBoxPosition();
            Invalidate();
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x0112;
            const int SC_MAXIMIZE = 0xF030;
            if (m.Msg == WM_SYSCOMMAND && (int)m.WParam == SC_MAXIMIZE)
            {
                ToggleFullscreen();
                return;
            }
            base.WndProc(ref m);
        }

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
                if (rotateBtnRect.Contains(x, y))
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
                    if (editMode)
                    {
                        SaveHomeworkData();
                    }
                    editMode = !editMode;
                    if (!editMode && editingSubjectIndex != -1)
                        FinishInlineEdit();
                    Invalidate();
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
                else if (editMode && !rotationMode)
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

                        if (fieldRects[i, 0] != Rectangle.Empty && fieldRects[i, 0].Contains(x, y))
                        {
                            StartInlineEdit(i, EditFieldType.DueTime, fieldRects[i, 0]);
                            return;
                        }
                    }
                }
            }
        }

        // 启动内联编辑
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

            if (fieldType == EditFieldType.Subject)
            {
                string currentText = homeworkData.Subjects.ContainsKey(subject) ? homeworkData.Subjects[subject] : "";
                var textBox = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Location = screenLoc,
                    Size = screenSize,
                    Text = currentText,
                    Font = font30,
                    BorderStyle = BorderStyle.FixedSingle
                };
                textBox.LostFocus += InlineEdit_LostFocus;
                textBox.KeyDown += InlineEdit_KeyDown;
                inlineEditControl = textBox;
            }
            else // DueTime
            {
                string currentValue = homeworkData.DueTimes.ContainsKey(subject) ? homeworkData.DueTimes[subject] : "";
                if (string.IsNullOrEmpty(currentValue))
                    currentValue = "晚修三";

                var comboBox = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = screenLoc,
                    Size = new Size(screenSize.Width, 25),
                    Font = font22,
                    FlatStyle = FlatStyle.Flat
                };
                comboBox.Items.AddRange(new object[] { "晚修一", "晚修二", "晚修三", "明天早读", "无" });
                comboBox.SelectedItem = currentValue;
                if (comboBox.SelectedItem == null)
                    comboBox.SelectedIndex = 2;

                comboBox.SelectedIndexChanged += InlineEdit_SelectedIndexChanged;
                comboBox.Leave += InlineEdit_LostFocus;
                comboBox.KeyDown += InlineEdit_KeyDown;
                comboBox.DroppedDown = true;
                inlineEditControl = comboBox;
            }

            this.Controls.Add(inlineEditControl);
            inlineEditControl.Focus();
        }

        private void InlineEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (currentEditType == EditFieldType.DueTime && sender is ComboBox combo && !combo.DroppedDown)
            {
                FinishInlineEdit();
            }
        }

        private void InlineEdit_LostFocus(object sender, EventArgs e)
        {
            FinishInlineEdit();
        }

        private void InlineEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (currentEditType == EditFieldType.DueTime && sender is ComboBox)
                {
                    e.SuppressKeyPress = true;
                    FinishInlineEdit();
                }
                else if (currentEditType == EditFieldType.Subject && e.Control)
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

            if (inlineEditControl is TextBox textBox)
            {
                textBox.LostFocus -= InlineEdit_LostFocus;
                textBox.KeyDown -= InlineEdit_KeyDown;
            }
            else if (inlineEditControl is ComboBox comboBox)
            {
                comboBox.SelectedIndexChanged -= InlineEdit_SelectedIndexChanged;
                comboBox.Leave -= InlineEdit_LostFocus;
                comboBox.KeyDown -= InlineEdit_KeyDown;
            }

            string subject = currentSubjects[editingSubjectIndex];
            string newValue = null;

            switch (currentEditType)
            {
                case EditFieldType.Subject:
                    newValue = ((TextBox)inlineEditControl).Text;
                    homeworkData.Subjects[subject] = newValue;
                    break;
                case EditFieldType.DueTime:
                    newValue = ((ComboBox)inlineEditControl).SelectedItem?.ToString() ?? "";
                    homeworkData.DueTimes[subject] = newValue;
                    break;
            }

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

            if (inlineEditControl is TextBox textBox)
            {
                textBox.LostFocus -= InlineEdit_LostFocus;
                textBox.KeyDown -= InlineEdit_KeyDown;
            }
            else if (inlineEditControl is ComboBox comboBox)
            {
                comboBox.SelectedIndexChanged -= InlineEdit_SelectedIndexChanged;
                comboBox.Leave -= InlineEdit_LostFocus;
                comboBox.KeyDown -= InlineEdit_KeyDown;
            }

            this.Controls.Remove(inlineEditControl);
            inlineEditControl.Dispose();
            inlineEditControl = null;
            editingSubjectIndex = -1;
            currentEditType = EditFieldType.None;
            if (!this.IsDisposed)
                Invalidate();
        }

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

        private void LoadImages()
        {
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Slide");
            bgImage = LoadImage(Path.Combine(imagePath, "背景.png"), VIRTUAL_SIZE);
            dialogBgImage = LoadImage(Path.Combine(imagePath, "对话框背景.png"), new Size(400, 200));
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

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            g.TranslateTransform(offset.X, offset.Y);
            g.ScaleTransform(scaleFactor, scaleFactor);

            if (bgImage != null)
                g.DrawImage(bgImage, new Rectangle(0, 0, VIRTUAL_SIZE.Width, VIRTUAL_SIZE.Height));
            else
            {
                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(200, 220, 255)))
                    g.FillRectangle(bgBrush, new Rectangle(0, 0, VIRTUAL_SIZE.Width, VIRTUAL_SIZE.Height));
            }

            DrawDateInfo(g);
            if (rotationMode)
                DrawRotationView(g);
            else
                DrawGrid(g);
            DrawButtons(g);

            g.ResetTransform();
        }

        private void DrawDateInfo(Graphics g)
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
                    int weekdayHeight = font20.Height;
                    int dateHeight = font12.Height;
                    int totalHeight = weekdayHeight + dateHeight;
                    int startY = (47 - totalHeight) / 2;

                    SizeF weekdaySize = g.MeasureString(weekdayText, font20);
                    g.DrawString(weekdayText, font20, whiteBrush, new PointF((VIRTUAL_SIZE.Width - weekdaySize.Width) / 2, startY));

                    SizeF dateSize = g.MeasureString(dateText, font12);
                    g.DrawString(dateText, font12, whiteBrush, new PointF((VIRTUAL_SIZE.Width - dateSize.Width) / 2, startY + weekdayHeight));
                }
                else
                {
                    g.DrawString(weekdayText, font20, whiteBrush, new PointF(20, 2));
                    g.DrawString(dateText, font12, whiteBrush, new PointF(20, 2 + font20.Height));
                }
            }
        }

        private void DrawButtons(Graphics g)
        {
            if (rotationMode)
            {
                int backY = (47 - 37) / 2;
                int backWidth = (backBtnImage != null) ? backBtnImage.Width : 80;
                backBtnRect = new Rectangle(20, backY, backWidth, 37);

                if (backBtnImage != null)
                    g.DrawImage(backBtnImage, backBtnRect);
                else
                {
                    using (SolidBrush btnBrush = new SolidBrush(Color.SteelBlue))
                        g.FillRectangle(btnBrush, backBtnRect);
                    g.DrawRectangle(Pens.Black, backBtnRect);
                    g.DrawString("返回", buttonFont, Brushes.White, backBtnRect, CenterStringFormat());
                }

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
                rotateBtnRect = new Rectangle(ROTATE_BTN_POS, BTN_SIZE);
                DrawButton(g, rotateBtnRect, "轮换");

                editBtnRect = new Rectangle(EDIT_BTN_POS, BTN_SIZE);
                DrawButton(g, editBtnRect, editMode ? "完成" : "编辑");

                historyBtnRect = new Rectangle(HISTORY_BTN_POS, BTN_SIZE);
                DrawButton(g, historyBtnRect, historyMode ? "返回" : "记录");

                fullscreenBtnRect = new Rectangle(FULLSCREEN_BTN_POS, BTN_SIZE);
                DrawButton(g, fullscreenBtnRect, fullscreen ? "缩小" : "全屏");
            }
        }

        private void DrawButton(Graphics g, Rectangle rect, string text)
        {
            if (buttonImage != null)
                g.DrawImage(buttonImage, rect);
            else
            {
                using (SolidBrush btnBrush = new SolidBrush(Color.SteelBlue))
                    g.FillRectangle(btnBrush, rect);
                g.DrawRectangle(Pens.Black, rect);
            }
            g.DrawString(text, buttonFont, Brushes.White, rect, CenterStringFormat());
        }

        private StringFormat CenterStringFormat()
        {
            return new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        }

        private void DrawGrid(Graphics g)
        {
            for (int i = 0; i < gridRects.Count; i++)
            {
                Rectangle rect = gridRects[i];
                using (LinearGradientBrush brush = new LinearGradientBrush(rect, Color.FromArgb(220, 220, 220), Color.White, LinearGradientMode.Vertical))
                {
                    g.FillRectangle(brush, rect);
                }
                g.DrawRectangle(new Pen(GRAY, GRID_BORDER), rect);

                if (i < currentSubjects.Length)
                {
                    string subject = currentSubjects[i];

                    g.DrawString(subject, font22, Brushes.Gray, new PointF(rect.Left + 10, rect.Top + 10));

                    int lineY = rect.Top + 50;
                    g.DrawLine(Pens.Gray, rect.Left + 10, lineY, rect.Right - 10, lineY);

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

                    g.DrawString(prefix, fontSmall, Brushes.Gray, startX, prefixY);
                    g.DrawString(displayTime, font22, Brushes.Red, startX + prefixWidth, timeY);

                    int minY = Math.Min(prefixY, timeY);
                    int maxY = lineY;
                    Rectangle timeRect = new Rectangle((int)startX, minY, (int)totalWidth + 5, maxY - minY + 2);
                    fieldRects[i, 0] = timeRect;

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
                        using (SolidBrush lightBrush = new SolidBrush(Color.FromArgb(240, 240, 240)))
                        {
                            g.FillRectangle(lightBrush, textArea);
                        }
                        g.DrawRectangle(Pens.LightGray, textArea);
                        string hint = editMode ? "点我编辑作业" : "今天暂时没有此项作业";
                        g.DrawString(hint, hintFont, Brushes.Gray, textArea, CenterStringFormat());
                    }
                }
            }
        }

        private void DrawRotationView(Graphics g)
        {
            Rectangle bigRect = new Rectangle(150, 100, 900, 475);
            using (LinearGradientBrush brush = new LinearGradientBrush(bigRect, Color.FromArgb(220, 220, 220), Color.White, LinearGradientMode.Vertical))
            {
                g.FillRectangle(brush, bigRect);
            }
            g.DrawRectangle(new Pen(GRAY, GRID_BORDER), bigRect);

            string subject = currentSubjects[rotationIndex];
            g.DrawString(subject, font36, Brushes.Gray, new PointF(600 - (int)g.MeasureString(subject, font36).Width / 2, bigRect.Top + 30));

            int lineY = bigRect.Top + 120;
            g.DrawLine(new Pen(GRAY, 2), bigRect.Left + 50, lineY, bigRect.Right - 50, lineY);

            Rectangle textArea = new Rectangle(bigRect.Left + 50, bigRect.Top + 150, bigRect.Width - 100, bigRect.Height - 200);
            if (homeworkData.Subjects.ContainsKey(subject) && !string.IsNullOrWhiteSpace(homeworkData.Subjects[subject]))
            {
                DrawTextInArea(g, homeworkData.Subjects[subject], textArea, font30, center: true);
            }
            else
            {
                g.DrawString("今天暂时没有此项作业", font30, Brushes.Gray, textArea, CenterStringFormat());
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
                    g.DrawString(lines[i], font, Brushes.Black, new PointF(area.Left + (area.Width - sz.Width) / 2, y));
                }
                else
                {
                    g.DrawString(lines[i], font, Brushes.Black, new PointF(area.Left + 5, y));
                }
            }
            if (lines.Count > maxLines)
            {
                g.DrawString("...", font, Brushes.Black, new PointF(area.Left + 5, area.Top + maxLines * lineHeight));
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            rotationTimer?.Dispose();
            base.OnFormClosed(e);
        }
    }
}