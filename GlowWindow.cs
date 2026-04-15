using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace HomeworkViewer
{
    public class GlowWindow : Form
    {
        private Point _mousePos = new Point(-1000, -1000);
        private System.Windows.Forms.Timer _animationTimer;
        private float _alpha = 0f;
        private bool _fadeIn = false;
        private bool _fadeOut = false;
        private int _radius = 120;

        public GlowWindow()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Black;
            TransparencyKey = Color.Black;
            Opacity = 0;
            DoubleBuffered = true;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);

            // 设置鼠标穿透
            int exStyle = (int)NativeMethods.GetWindowLong(this.Handle, NativeMethods.GWL_EXSTYLE);
            exStyle |= NativeMethods.WS_EX_TRANSPARENT | NativeMethods.WS_EX_LAYERED | NativeMethods.WS_EX_TOOLWINDOW;
            NativeMethods.SetWindowLong(this.Handle, NativeMethods.GWL_EXSTYLE, (IntPtr)exStyle);

            _animationTimer = new System.Windows.Forms.Timer { Interval = 16 };
            _animationTimer.Tick += (s, e) => UpdateAlpha();
            _animationTimer.Start();
        }

        private void UpdateAlpha()
        {
            if (_fadeIn)
            {
                _alpha += 0.05f;
                if (_alpha >= 0.8f) { _alpha = 0.8f; _fadeIn = false; }
                Opacity = _alpha;
                Invalidate();
            }
            else if (_fadeOut)
            {
                _alpha -= 0.05f;
                if (_alpha <= 0f) { _alpha = 0f; _fadeOut = false; Opacity = 0; }
                else Opacity = _alpha;
                Invalidate();
            }
        }

        public void SetPosition(Point screenPos)
        {
            _mousePos = screenPos;
            int x = screenPos.X - _radius;
            int y = screenPos.Y - _radius;
            Location = new Point(x, y);
            Size = new Size(_radius * 2, _radius * 2);
        }

        public void StartFadeIn()
        {
            _fadeIn = true;
            _fadeOut = false;
        }

        public void StartFadeOut()
        {
            _fadeIn = false;
            _fadeOut = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (_alpha <= 0) return;

            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            using (var path = new GraphicsPath())
            {
                path.AddEllipse(0, 0, _radius * 2, _radius * 2);
                using (var brush = new PathGradientBrush(path))
                {
                    brush.CenterColor = Color.FromArgb((int)(_alpha * 80), 255, 255, 255);
                    brush.SurroundColors = new Color[] { Color.FromArgb(0, 255, 255, 255) };
                    brush.CenterPoint = new PointF(_radius, _radius);
                    g.FillPath(brush, path);
                }
            }
        }

        private static class NativeMethods
        {
            public const int GWL_EXSTYLE = -20;
            public const int WS_EX_TRANSPARENT = 0x20;
            public const int WS_EX_LAYERED = 0x80000;
            public const int WS_EX_TOOLWINDOW = 0x80;

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            public static extern IntPtr GetWindowLong(IntPtr hWnd, int nIndex);

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            public static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        }
    }
}