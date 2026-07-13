using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace ToolrackCapture
{
    public static class Native
    {
        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out NativeRect rect);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetShellWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(Point point, uint flags);

        [DllImport("shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr monitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(
            IntPtr hwnd,
            IntPtr insertAfter,
            int x,
            int y,
            int width,
            int height,
            uint flags);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(
            IntPtr hwnd,
            int attribute,
            out NativeRect value,
            int size);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(
            IntPtr hwnd,
            int attribute,
            ref int value,
            int size);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(
            IntPtr hwnd,
            int attribute,
            out int value,
            int size);

        public static void EnableDpiAwareness()
        {
            try
            {
                if (SetProcessDpiAwarenessContext(new IntPtr(-4)))
                {
                    return;
                }
            }
            catch (EntryPointNotFoundException)
            {
            }
            catch (DllNotFoundException)
            {
            }

            try
            {
                SetProcessDPIAware();
            }
            catch
            {
            }
        }

        public static Bitmap CaptureRectangle(Rectangle rectangle)
        {
            if (rectangle.Width < 1 || rectangle.Height < 1)
            {
                throw new ArgumentOutOfRangeException("rectangle");
            }

            Bitmap bitmap = new Bitmap(
                rectangle.Width,
                rectangle.Height,
                PixelFormat.Format32bppPArgb);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(
                    rectangle.X,
                    rectangle.Y,
                    0,
                    0,
                    rectangle.Size,
                    CopyPixelOperation.SourceCopy);
            }
            return bitmap;
        }

        public static Rectangle GetWindowRectangle(IntPtr hwnd)
        {
            NativeRect nativeRect;
            if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out nativeRect))
            {
                return Rectangle.Empty;
            }
            return Rectangle.FromLTRB(
                nativeRect.Left,
                nativeRect.Top,
                nativeRect.Right,
                nativeRect.Bottom);
        }

        public static void PositionWindow(IntPtr hwnd, int x, int y)
        {
            const uint noSize = 0x0001;
            const uint noZOrder = 0x0004;
            const uint noActivate = 0x0010;
            if (!SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0, noSize | noZOrder | noActivate))
            {
                throw new InvalidOperationException("Could not position the capture palette.");
            }
        }

        public static Rectangle FindWindowRectangle(Point point)
        {
            Rectangle found = Rectangle.Empty;
            uint ownProcessId = (uint)Process.GetCurrentProcess().Id;
            IntPtr shellWindow = GetShellWindow();

            EnumWindows(delegate(IntPtr hwnd, IntPtr lParam)
            {
                if (hwnd == shellWindow || !IsWindowVisible(hwnd) || IsIconic(hwnd))
                {
                    return true;
                }

                uint processId;
                GetWindowThreadProcessId(hwnd, out processId);
                if (processId == ownProcessId)
                {
                    return true;
                }

                int cloaked = 0;
                try
                {
                    if (DwmGetWindowAttribute(hwnd, 14, out cloaked, sizeof(int)) == 0 && cloaked != 0)
                    {
                        return true;
                    }
                }
                catch
                {
                }

                NativeRect nativeRect = new NativeRect();
                bool hasRect = false;
                try
                {
                    hasRect = DwmGetWindowAttribute(
                        hwnd,
                        9,
                        out nativeRect,
                        Marshal.SizeOf(typeof(NativeRect))) == 0;
                }
                catch
                {
                }

                if (!hasRect)
                {
                    hasRect = GetWindowRect(hwnd, out nativeRect);
                }
                if (!hasRect)
                {
                    return true;
                }

                Rectangle candidate = Rectangle.FromLTRB(
                    nativeRect.Left,
                    nativeRect.Top,
                    nativeRect.Right,
                    nativeRect.Bottom);
                if (candidate.Width < 2 || candidate.Height < 2 || !candidate.Contains(point))
                {
                    return true;
                }

                found = candidate;
                return false;
            }, IntPtr.Zero);

            return found;
        }

        public static string ApplyPaletteBackdrop(IntPtr hwnd, bool dark, bool forceFailure)
        {
            if (forceFailure || hwnd == IntPtr.Zero)
            {
                return "solid-fallback";
            }
            try
            {
                int darkValue = dark ? 1 : 0;
                int cornerPreference = 2;
                int borderColor = unchecked((int)0xFFFFFFFE);
                int darkResult = DwmSetWindowAttribute(hwnd, 20, ref darkValue, sizeof(int));
                int cornerResult = DwmSetWindowAttribute(hwnd, 33, ref cornerPreference, sizeof(int));
                int borderResult = DwmSetWindowAttribute(hwnd, 34, ref borderColor, sizeof(int));
                if (Environment.OSVersion.Version.Build >= 22621)
                {
                    int backdrop = 3;
                    int backdropResult = DwmSetWindowAttribute(hwnd, 38, ref backdrop, sizeof(int));
                    if (darkResult == 0 && cornerResult == 0 && borderResult == 0 && backdropResult == 0)
                    {
                        return "transient";
                    }
                }
                if (cornerResult == 0)
                {
                    return "rounded-solid";
                }
            }
            catch
            {
            }
            return "solid-fallback";
        }

        public static float GetDpiScaleAtPoint(Point point)
        {
            try
            {
                IntPtr monitor = MonitorFromPoint(point, 2);
                uint dpiX;
                uint dpiY;
                if (monitor != IntPtr.Zero && GetDpiForMonitor(monitor, 0, out dpiX, out dpiY) == 0 && dpiX >= 96)
                {
                    return dpiX / 96.0f;
                }
            }
            catch
            {
            }
            try
            {
                using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
                {
                    return Math.Max(1.0f, graphics.DpiX / 96.0f);
                }
            }
            catch
            {
                return 1.0f;
            }
        }
    }

    public static class BuildInfo
    {
        public const string Id = "capture-native-v1";
    }

    public abstract class SelectorForm : Form
    {
        protected Rectangle highlighted = Rectangle.Empty;

        protected SelectorForm(Cursor cursor)
        {
            AutoScaleMode = AutoScaleMode.None;
            BackColor = Color.FromArgb(20, 20, 23);
            Bounds = SystemInformation.VirtualScreen;
            Cursor = cursor;
            DoubleBuffered = true;
            FormBorderStyle = FormBorderStyle.None;
            KeyPreview = true;
            Opacity = 0.25d;
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (highlighted.Width < 1 || highlighted.Height < 1)
            {
                return;
            }

            Rectangle local = highlighted;
            local.Offset(-Bounds.Left, -Bounds.Top);
            local.Width = Math.Max(1, local.Width - 1);
            local.Height = Math.Max(1, local.Height - 1);
            using (Pen outer = new Pen(Color.White, 9.0f))
            using (Pen inner = new Pen(Color.FromArgb(80, 140, 255), 3.0f))
            {
                e.Graphics.DrawRectangle(outer, local);
                e.Graphics.DrawRectangle(inner, local);
            }
        }
    }

    public sealed class RangeSelector : SelectorForm
    {
        private Point start;
        private bool dragging;

        public Rectangle Selection { get; private set; }

        private RangeSelector()
            : base(Cursors.Cross)
        {
            Selection = Rectangle.Empty;
        }

        public static Rectangle SelectRectangle()
        {
            using (RangeSelector selector = new RangeSelector())
            {
                return selector.ShowDialog() == DialogResult.OK
                    ? selector.Selection
                    : Rectangle.Empty;
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                start = Control.MousePosition;
                dragging = true;
                highlighted = Rectangle.Empty;
                Invalidate();
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (dragging)
            {
                Point current = Control.MousePosition;
                highlighted = Rectangle.FromLTRB(
                    Math.Min(start.X, current.X),
                    Math.Min(start.Y, current.Y),
                    Math.Max(start.X, current.X),
                    Math.Max(start.Y, current.Y));
                Invalidate();
            }
            base.OnMouseMove(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (dragging && e.Button == MouseButtons.Left)
            {
                dragging = false;
                if (highlighted.Width >= 2 && highlighted.Height >= 2)
                {
                    Selection = highlighted;
                    DialogResult = DialogResult.OK;
                    Close();
                    return;
                }
                highlighted = Rectangle.Empty;
                Invalidate();
            }
            base.OnMouseUp(e);
        }
    }

    public sealed class WindowSelector : SelectorForm
    {
        public Rectangle Selection { get; private set; }

        private WindowSelector()
            : base(Cursors.Hand)
        {
            Selection = Rectangle.Empty;
        }

        public static Rectangle SelectRectangle()
        {
            using (WindowSelector selector = new WindowSelector())
            {
                return selector.ShowDialog() == DialogResult.OK
                    ? selector.Selection
                    : Rectangle.Empty;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            Rectangle next = Native.FindWindowRectangle(Control.MousePosition);
            if (next != highlighted)
            {
                highlighted = next;
                Invalidate();
            }
            base.OnMouseMove(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && highlighted.Width > 0 && highlighted.Height > 0)
            {
                Selection = highlighted;
                DialogResult = DialogResult.OK;
                Close();
                return;
            }
            base.OnMouseDown(e);
        }
    }

    public sealed class PaletteHitAudit
    {
        public string Id { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    public sealed class PaletteKeyboardAudit
    {
        public bool TabCycles { get; set; }
        public bool ArrowsNavigate { get; set; }
        public bool EnterActivates { get; set; }
        public bool EscapeCloses { get; set; }
    }

    public sealed class PaletteAudit
    {
        public int ButtonCount { get; set; }
        public string[] Columns { get; set; }
        public string[] RangeActions { get; set; }
        public string[] WindowActions { get; set; }
        public int InnerCardBorders { get; set; }
        public int DividerCount { get; set; }
        public bool IconsAreVector { get; set; }
        public bool UsesIconFont { get; set; }
        public double LightTextContrast { get; set; }
        public double DarkTextContrast { get; set; }
        public int PanelWidth { get; set; }
        public int PanelHeight { get; set; }
        public Rectangle PanelBounds { get; set; }
        public PaletteHitAudit[] HitRectangles { get; set; }
        public PaletteKeyboardAudit Keyboard { get; set; }
        public bool ClosesOnDeactivate { get; set; }
        public string DwmMode { get; set; }
    }

    internal sealed class PaletteTheme
    {
        internal Color Surface;
        internal Color SurfaceEdge;
        internal Color Text;
        internal Color Muted;
        internal Color Divider;
        internal Color Hover;
        internal Color Pressed;
        internal Color Accent;

        internal static PaletteTheme Create(bool dark)
        {
            if (dark)
            {
                return new PaletteTheme
                {
                    Surface = Color.FromArgb(32, 33, 36),
                    SurfaceEdge = Color.FromArgb(72, 74, 80),
                    Text = Color.FromArgb(247, 247, 249),
                    Muted = Color.FromArgb(181, 184, 193),
                    Divider = Color.FromArgb(62, 64, 70),
                    Hover = Color.FromArgb(48, 50, 55),
                    Pressed = Color.FromArgb(57, 60, 67),
                    Accent = Color.FromArgb(129, 166, 255)
                };
            }
            return new PaletteTheme
            {
                Surface = Color.FromArgb(249, 249, 247),
                SurfaceEdge = Color.FromArgb(217, 218, 221),
                Text = Color.FromArgb(29, 29, 31),
                Muted = Color.FromArgb(92, 94, 101),
                Divider = Color.FromArgb(224, 224, 221),
                Hover = Color.FromArgb(238, 239, 238),
                Pressed = Color.FromArgb(227, 230, 233),
                Accent = Color.FromArgb(55, 103, 214)
            };
        }
    }

    internal sealed class PaletteLayout
    {
        internal const int LogicalWidth = 388;
        internal const int LogicalHeight = 230;
        internal float Scale;
        internal int Width;
        internal int Height;
        internal int DividerX;
        internal Rectangle PanelBounds;
        internal Rectangle[] Hits;
        internal string[] HitIds;

        internal static PaletteLayout Create(float requestedScale, Rectangle workArea, Point cursor)
        {
            float scale = Math.Max(1.0f, Math.Min(2.5f, requestedScale));
            PaletteLayout result = new PaletteLayout();
            result.Scale = scale;
            result.Width = Pixel(LogicalWidth, scale);
            result.Height = Pixel(LogicalHeight, scale);
            result.DividerX = result.Width / 2;
            int margin = Pixel(12, scale);
            int centerGap = Pixel(16, scale);
            int headerBottom = Pixel(50, scale);
            int rowHeight = Pixel(48, scale);
            int rowGap = Pixel(4, scale);
            int columnWidth = result.DividerX - margin - (centerGap / 2);
            int leftX = margin;
            int rightX = result.DividerX + (centerGap / 2);
            result.Hits = new Rectangle[6];
            result.HitIds = new[] {
                "range:image", "window:image", "range:path", "window:path", "range:text", "window:text"
            };
            for (int row = 0; row < 3; row++)
            {
                int y = headerBottom + (row * (rowHeight + rowGap));
                result.Hits[row * 2] = new Rectangle(leftX, y, columnWidth, rowHeight);
                result.Hits[(row * 2) + 1] = new Rectangle(rightX, y, columnWidth, rowHeight);
            }

            int offset = Pixel(14, scale);
            int inset = Pixel(8, scale);
            int x = cursor.X + offset;
            int yPosition = cursor.Y + offset;
            if (x + result.Width > workArea.Right - inset) { x = cursor.X - result.Width - offset; }
            if (yPosition + result.Height > workArea.Bottom - inset) { yPosition = cursor.Y - result.Height - offset; }
            x = Math.Max(workArea.Left + inset, Math.Min(x, workArea.Right - inset - result.Width));
            yPosition = Math.Max(workArea.Top + inset, Math.Min(yPosition, workArea.Bottom - inset - result.Height));
            result.PanelBounds = new Rectangle(x, yPosition, result.Width, result.Height);
            return result;
        }

        internal static int Pixel(int logical, float scale)
        {
            return Math.Max(1, (int)Math.Round(logical * scale, MidpointRounding.AwayFromZero));
        }
    }

    internal sealed class PaletteFocusMachine
    {
        internal int Index { get; private set; }
        internal bool EscapeRequested { get; private set; }

        internal PaletteFocusMachine()
        {
            Index = 0;
        }

        internal void Tab(bool reverse)
        {
            Index = reverse ? (Index + 5) % 6 : (Index + 1) % 6;
        }

        internal void Move(Keys key)
        {
            int row = Index / 2;
            int column = Index % 2;
            if (key == Keys.Left) { column = 0; }
            else if (key == Keys.Right) { column = 1; }
            else if (key == Keys.Up) { row = (row + 2) % 3; }
            else if (key == Keys.Down) { row = (row + 1) % 3; }
            Index = (row * 2) + column;
        }

        internal string Activate(PaletteLayout layout)
        {
            return layout.HitIds[Index];
        }

        internal void Escape()
        {
            EscapeRequested = true;
        }

        internal static PaletteKeyboardAudit Audit()
        {
            PaletteLayout layout = PaletteLayout.Create(1.0f, new Rectangle(0, 0, 1000, 800), new Point(100, 100));
            PaletteFocusMachine tab = new PaletteFocusMachine();
            for (int index = 0; index < 6; index++) { tab.Tab(false); }
            bool tabCycles = tab.Index == 0;

            PaletteFocusMachine arrows = new PaletteFocusMachine();
            arrows.Move(Keys.Right);
            bool right = arrows.Index == 1;
            arrows.Move(Keys.Down);
            bool down = arrows.Index == 3;
            arrows.Move(Keys.Left);
            bool left = arrows.Index == 2;
            arrows.Move(Keys.Up);
            bool up = arrows.Index == 0;
            string activated = arrows.Activate(layout);
            arrows.Escape();
            return new PaletteKeyboardAudit
            {
                TabCycles = tabCycles,
                ArrowsNavigate = right && down && left && up,
                EnterActivates = String.Equals(activated, "range:image", StringComparison.Ordinal),
                EscapeCloses = arrows.EscapeRequested
            };
        }
    }

    internal static class PaletteRenderer
    {
        private static readonly string[] ActionLabels = {
            "\u753b\u50cf", "\u753b\u50cf", "\u30d1\u30b9", "\u30d1\u30b9", "\u30c6\u30ad\u30b9\u30c8", "\u30c6\u30ad\u30b9\u30c8"
        };

        internal static GraphicsPath RoundedRectangle(Rectangle rectangle, float radius)
        {
            GraphicsPath path = new GraphicsPath();
            float diameter = Math.Max(1.0f, radius * 2.0f);
            RectangleF arc = new RectangleF(rectangle.X, rectangle.Y, diameter, diameter);
            path.AddArc(arc, 180, 90);
            arc.X = rectangle.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = rectangle.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = rectangle.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }

        internal static void Draw(Graphics graphics, PaletteLayout layout, PaletteTheme theme, int hover, int pressed, int focus)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            graphics.Clear(Color.Transparent);
            Rectangle surface = new Rectangle(0, 0, layout.Width - 1, layout.Height - 1);
            using (GraphicsPath outer = RoundedRectangle(surface, 18.0f * layout.Scale))
            using (SolidBrush surfaceBrush = new SolidBrush(theme.Surface))
            using (Pen edge = new Pen(theme.SurfaceEdge, Math.Max(1.0f, layout.Scale)))
            {
                graphics.FillPath(surfaceBrush, outer);
                graphics.DrawPath(edge, outer);
            }

            int dividerTop = PaletteLayout.Pixel(18, layout.Scale);
            int dividerBottom = layout.Height - PaletteLayout.Pixel(18, layout.Scale);
            using (Pen divider = new Pen(theme.Divider, Math.Max(1.0f, layout.Scale)))
            {
                graphics.DrawLine(divider, layout.DividerX, dividerTop, layout.DividerX, dividerBottom);
            }

            DrawHeader(graphics, layout, theme, false);
            DrawHeader(graphics, layout, theme, true);
            for (int index = 0; index < layout.Hits.Length; index++)
            {
                Rectangle hit = layout.Hits[index];
                if (index == hover || index == pressed)
                {
                    using (GraphicsPath hoverPath = RoundedRectangle(hit, 11.0f * layout.Scale))
                    using (SolidBrush brush = new SolidBrush(index == pressed ? theme.Pressed : theme.Hover))
                    {
                        graphics.FillPath(brush, hoverPath);
                    }
                }
                if (index == focus)
                {
                    Rectangle focusRect = hit;
                    focusRect.Inflate(-PaletteLayout.Pixel(2, layout.Scale), -PaletteLayout.Pixel(2, layout.Scale));
                    using (GraphicsPath focusPath = RoundedRectangle(focusRect, 10.0f * layout.Scale))
                    using (Pen focusPen = new Pen(theme.Accent, Math.Max(1.5f, 1.5f * layout.Scale)))
                    {
                        graphics.DrawPath(focusPen, focusPath);
                    }
                }
                DrawAction(graphics, layout, theme, index, hit);
            }
        }

        private static void DrawHeader(Graphics graphics, PaletteLayout layout, PaletteTheme theme, bool window)
        {
            int startX = window ? layout.DividerX + PaletteLayout.Pixel(19, layout.Scale) : PaletteLayout.Pixel(19, layout.Scale);
            Rectangle icon = new Rectangle(startX, PaletteLayout.Pixel(16, layout.Scale), PaletteLayout.Pixel(22, layout.Scale), PaletteLayout.Pixel(18, layout.Scale));
            DrawHeaderIcon(graphics, icon, theme.Accent, window, layout.Scale);
            string label = window ? "\u30a6\u30a3\u30f3\u30c9\u30a6" : "\u7bc4\u56f2";
            using (Font font = new Font("Segoe UI", 10.5f * layout.Scale, FontStyle.Bold, GraphicsUnit.Pixel))
            using (SolidBrush brush = new SolidBrush(theme.Text))
            using (StringFormat format = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter })
            {
                Rectangle text = new Rectangle(icon.Right + PaletteLayout.Pixel(8, layout.Scale), icon.Top - PaletteLayout.Pixel(2, layout.Scale),
                    PaletteLayout.Pixel(126, layout.Scale), icon.Height + PaletteLayout.Pixel(4, layout.Scale));
                graphics.DrawString(label, font, brush, text, format);
            }
        }

        private static void DrawHeaderIcon(Graphics graphics, Rectangle box, Color color, bool window, float scale)
        {
            using (GraphicsPath path = new GraphicsPath())
            using (Pen pen = new Pen(color, Math.Max(1.4f, 1.5f * scale)))
            {
                path.AddRectangle(new RectangleF(box.X + 1, box.Y + 1, box.Width - 3, box.Height - 3));
                if (window)
                {
                    path.StartFigure();
                    path.AddLine(box.X + 2, box.Y + (box.Height * 0.34f), box.Right - 3, box.Y + (box.Height * 0.34f));
                }
                else
                {
                    pen.DashStyle = DashStyle.Dash;
                }
                graphics.DrawPath(pen, path);
            }
        }

        private static void DrawAction(Graphics graphics, PaletteLayout layout, PaletteTheme theme, int index, Rectangle hit)
        {
            int iconSize = PaletteLayout.Pixel(22, layout.Scale);
            Rectangle icon = new Rectangle(hit.X + PaletteLayout.Pixel(15, layout.Scale), hit.Y + ((hit.Height - iconSize) / 2), iconSize, iconSize);
            int actionRow = index / 2;
            DrawActionIcon(graphics, icon, actionRow, theme.Muted, layout.Scale);
            using (Font font = new Font("Segoe UI", 11.5f * layout.Scale, FontStyle.Regular, GraphicsUnit.Pixel))
            using (SolidBrush brush = new SolidBrush(theme.Text))
            using (StringFormat format = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter })
            {
                Rectangle text = new Rectangle(icon.Right + PaletteLayout.Pixel(11, layout.Scale), hit.Y,
                    hit.Right - icon.Right - PaletteLayout.Pixel(17, layout.Scale), hit.Height);
                graphics.DrawString(ActionLabels[index], font, brush, text, format);
            }
        }

        private static void DrawActionIcon(Graphics graphics, Rectangle box, int action, Color color, float scale)
        {
            using (GraphicsPath path = new GraphicsPath())
            using (Pen pen = new Pen(color, Math.Max(1.35f, 1.45f * scale)))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;
                if (action == 0)
                {
                    path.AddRectangle(new RectangleF(box.X + 2, box.Y + 3, box.Width - 5, box.Height - 7));
                    path.StartFigure();
                    path.AddEllipse(box.X + box.Width * 0.62f, box.Y + box.Height * 0.25f, box.Width * 0.14f, box.Height * 0.14f);
                    path.StartFigure();
                    path.AddLines(new[] {
                        new PointF(box.X + box.Width * 0.16f, box.Y + box.Height * 0.72f),
                        new PointF(box.X + box.Width * 0.42f, box.Y + box.Height * 0.48f),
                        new PointF(box.X + box.Width * 0.57f, box.Y + box.Height * 0.63f),
                        new PointF(box.X + box.Width * 0.72f, box.Y + box.Height * 0.48f),
                        new PointF(box.X + box.Width * 0.84f, box.Y + box.Height * 0.63f)
                    });
                }
                else if (action == 1)
                {
                    path.AddArc(box.X + 2, box.Y + box.Height * 0.34f, box.Width * 0.52f, box.Height * 0.46f, 135, 270);
                    path.StartFigure();
                    path.AddArc(box.X + box.Width * 0.42f, box.Y + box.Height * 0.20f, box.Width * 0.52f, box.Height * 0.46f, -45, 270);
                    path.StartFigure();
                    path.AddLine(box.X + box.Width * 0.35f, box.Y + box.Height * 0.65f, box.X + box.Width * 0.67f, box.Y + box.Height * 0.35f);
                }
                else
                {
                    path.AddLine(box.X + box.Width * 0.18f, box.Y + box.Height * 0.22f, box.X + box.Width * 0.82f, box.Y + box.Height * 0.22f);
                    path.StartFigure();
                    path.AddLine(box.X + box.Width * 0.50f, box.Y + box.Height * 0.22f, box.X + box.Width * 0.50f, box.Y + box.Height * 0.80f);
                    path.StartFigure();
                    path.AddLine(box.X + box.Width * 0.34f, box.Y + box.Height * 0.80f, box.X + box.Width * 0.66f, box.Y + box.Height * 0.80f);
                }
                graphics.DrawPath(pen, path);
            }
        }
    }

    internal sealed class PaletteSurface : Control
    {
        private readonly PaletteLayout layout;
        private readonly PaletteTheme theme;
        private readonly Action<int> choose;
        private int hover = -1;
        private int pressed = -1;
        private int focus;
        private bool markerWritten;

        internal PaletteSurface(PaletteLayout paletteLayout, PaletteTheme paletteTheme, Action<int> choice)
        {
            layout = paletteLayout;
            theme = paletteTheme;
            choose = choice;
            Dock = DockStyle.Fill;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
        }

        internal int FocusIndex
        {
            get { return focus; }
            set { focus = Math.Max(0, Math.Min(5, value)); Invalidate(); }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            PaletteRenderer.Draw(e.Graphics, layout, theme, hover, pressed, focus);
            WritePaintMarker();
            base.OnPaint(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int next = HitTest(e.Location);
            if (next != hover) { hover = next; Invalidate(); }
            base.OnMouseMove(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            hover = -1;
            pressed = -1;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                pressed = HitTest(e.Location);
                if (pressed >= 0) { focus = pressed; Focus(); }
                Invalidate();
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            int released = HitTest(e.Location);
            int selected = pressed;
            pressed = -1;
            Invalidate();
            if (e.Button == MouseButtons.Left && selected >= 0 && released == selected) { choose(selected); }
            base.OnMouseUp(e);
        }

        private int HitTest(Point point)
        {
            for (int index = 0; index < layout.Hits.Length; index++)
            {
                if (layout.Hits[index].Contains(point)) { return index; }
            }
            return -1;
        }

        private void WritePaintMarker()
        {
            if (markerWritten) { return; }
            markerWritten = true;
            string directory = Environment.GetEnvironmentVariable("TOOLRACK_CAPTURE_PAINT_DIR");
            if (!String.IsNullOrWhiteSpace(directory))
            {
                try
                {
                    Directory.CreateDirectory(directory);
                    string path = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".tick");
                    File.WriteAllText(path, Stopwatch.GetTimestamp().ToString(), new UTF8Encoding(false));
                }
                catch
                {
                }
            }
            if (String.Equals(Environment.GetEnvironmentVariable("TOOLRACK_CAPTURE_AUTOCLOSE"), "1", StringComparison.Ordinal))
            {
                BeginInvoke((MethodInvoker)delegate { FindForm().Close(); });
            }
        }
    }

    internal sealed class CapturePaletteForm : Form
    {
        private readonly PaletteLayout layout;
        private readonly PaletteFocusMachine focus = new PaletteFocusMachine();
        private readonly PaletteSurface surface;
        private bool ready;

        internal string SelectedAction { get; private set; }

        internal CapturePaletteForm(Point cursor, Rectangle workArea, float scale, bool dark)
        {
            SelectedAction = String.Empty;
            layout = PaletteLayout.Create(scale, workArea, cursor);
            PaletteTheme theme = PaletteTheme.Create(dark);
            AutoScaleMode = AutoScaleMode.None;
            BackColor = theme.Surface;
            Bounds = layout.PanelBounds;
            ClientSize = new Size(layout.Width, layout.Height);
            DoubleBuffered = true;
            FormBorderStyle = FormBorderStyle.None;
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            surface = new PaletteSurface(layout, theme, Choose);
            Controls.Add(surface);
            using (GraphicsPath regionPath = PaletteRenderer.RoundedRectangle(new Rectangle(0, 0, layout.Width, layout.Height), 18.0f * scale))
            {
                Region = new Region(regionPath);
            }
            if (String.Equals(Environment.GetEnvironmentVariable("TOOLRACK_CAPTURE_TEST_OFFSCREEN"), "1", StringComparison.Ordinal))
            {
                Opacity = 0.01d;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Native.ApplyPaletteBackdrop(Handle, ThemeDetector.IsDark(), false);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            ready = true;
            Activate();
            surface.Focus();
        }

        protected override void OnDeactivate(EventArgs e)
        {
            if (ready && String.IsNullOrEmpty(SelectedAction)) { Close(); }
            base.OnDeactivate(e);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys key = keyData & Keys.KeyCode;
            if (key == Keys.Tab)
            {
                focus.Tab((keyData & Keys.Shift) == Keys.Shift);
                surface.FocusIndex = focus.Index;
                return true;
            }
            if (key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down)
            {
                focus.Move(key);
                surface.FocusIndex = focus.Index;
                return true;
            }
            if (key == Keys.Enter || key == Keys.Space)
            {
                SelectedAction = focus.Activate(layout);
                DialogResult = DialogResult.OK;
                Close();
                return true;
            }
            if (key == Keys.Escape)
            {
                focus.Escape();
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void Choose(int index)
        {
            surface.FocusIndex = index;
            SelectedAction = layout.HitIds[index];
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    internal static class ThemeDetector
    {
        internal static bool IsDark()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"))
                {
                    if (key == null) { return false; }
                    object value = key.GetValue("AppsUseLightTheme");
                    return value != null && Convert.ToInt32(value) == 0;
                }
            }
            catch
            {
                return false;
            }
        }
    }

    public static class CapturePalette
    {
        public static string ShowPalette()
        {
            Native.EnableDpiAwareness();
            Point cursor = Control.MousePosition;
            Screen screen = Screen.FromPoint(cursor);
            float scale = Native.GetDpiScaleAtPoint(cursor);
            using (CapturePaletteForm palette = new CapturePaletteForm(cursor, screen.WorkingArea, scale, ThemeDetector.IsDark()))
            {
                palette.ShowDialog();
                return palette.SelectedAction;
            }
        }

        public static PaletteAudit CreateAudit(double requestedScale, bool dark, Rectangle workArea, Point cursor, bool forceDwmFailure)
        {
            PaletteLayout layout = PaletteLayout.Create((float)requestedScale, workArea, cursor);
            PaletteHitAudit[] hits = new PaletteHitAudit[layout.Hits.Length];
            for (int index = 0; index < hits.Length; index++)
            {
                Rectangle rectangle = layout.Hits[index];
                hits[index] = new PaletteHitAudit
                {
                    Id = layout.HitIds[index], X = rectangle.X, Y = rectangle.Y,
                    Width = rectangle.Width, Height = rectangle.Height
                };
            }
            PaletteTheme light = PaletteTheme.Create(false);
            PaletteTheme darkTheme = PaletteTheme.Create(true);
            return new PaletteAudit
            {
                ButtonCount = 6,
                Columns = new[] { "range", "window" },
                RangeActions = new[] { "image", "path", "text" },
                WindowActions = new[] { "image", "path", "text" },
                InnerCardBorders = 0,
                DividerCount = 1,
                IconsAreVector = true,
                UsesIconFont = false,
                LightTextContrast = Contrast(light.Text, light.Surface),
                DarkTextContrast = Contrast(darkTheme.Text, darkTheme.Surface),
                PanelWidth = layout.Width,
                PanelHeight = layout.Height,
                PanelBounds = layout.PanelBounds,
                HitRectangles = hits,
                Keyboard = PaletteFocusMachine.Audit(),
                ClosesOnDeactivate = true,
                DwmMode = Native.ApplyPaletteBackdrop(IntPtr.Zero, dark, forceDwmFailure)
            };
        }

        public static void RenderPreview(string path, bool dark, double requestedScale)
        {
            if (String.IsNullOrWhiteSpace(path)) { throw new ArgumentException("preview path is required", "path"); }
            float scale = (float)requestedScale;
            PaletteLayout layout = PaletteLayout.Create(scale, new Rectangle(0, 0, 4000, 3000), new Point(100, 100));
            PaletteTheme theme = PaletteTheme.Create(dark);
            string parent = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!Directory.Exists(parent)) { Directory.CreateDirectory(parent); }
            using (Bitmap bitmap = new Bitmap(layout.Width, layout.Height, PixelFormat.Format32bppPArgb))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Transparent);
                PaletteRenderer.Draw(graphics, layout, theme, -1, -1, 0);
                bitmap.Save(path, ImageFormat.Png);
            }
        }

        private static double Contrast(Color first, Color second)
        {
            double firstValue = Luminance(first);
            double secondValue = Luminance(second);
            double light = Math.Max(firstValue, secondValue);
            double dark = Math.Min(firstValue, secondValue);
            return (light + 0.05d) / (dark + 0.05d);
        }

        private static double Luminance(Color color)
        {
            return (0.2126d * Channel(color.R)) + (0.7152d * Channel(color.G)) + (0.0722d * Channel(color.B));
        }

        private static double Channel(byte value)
        {
            double normalized = value / 255.0d;
            return normalized <= 0.03928d ? normalized / 12.92d : Math.Pow((normalized + 0.055d) / 1.055d, 2.4d);
        }
    }

    internal sealed class ToastForm : Form
    {
        private readonly string message;
        private readonly PaletteTheme theme;
        private readonly Timer timer;

        internal ToastForm(string text)
        {
            message = text;
            theme = PaletteTheme.Create(ThemeDetector.IsDark());
            AutoScaleMode = AutoScaleMode.None;
            BackColor = theme.Surface;
            ClientSize = new Size(310, 58);
            DoubleBuffered = true;
            FormBorderStyle = FormBorderStyle.None;
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            Point cursor = Control.MousePosition;
            Rectangle work = Screen.FromPoint(cursor).WorkingArea;
            int x = Math.Min(work.Right - Width - 10, Math.Max(work.Left + 10, cursor.X + 14));
            int y = Math.Min(work.Bottom - Height - 10, Math.Max(work.Top + 10, cursor.Y + 14));
            Location = new Point(x, y);
            using (GraphicsPath path = PaletteRenderer.RoundedRectangle(new Rectangle(0, 0, Width, Height), 15.0f))
            {
                Region = new Region(path);
            }
            timer = new Timer();
            timer.Interval = 850;
            timer.Tick += delegate { timer.Stop(); Close(); };
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            timer.Start();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (SolidBrush background = new SolidBrush(theme.Surface))
            using (SolidBrush accent = new SolidBrush(theme.Accent))
            using (SolidBrush text = new SolidBrush(theme.Text))
            using (Font font = new Font("Segoe UI", 12.0f, FontStyle.Regular, GraphicsUnit.Pixel))
            using (Pen edge = new Pen(theme.SurfaceEdge))
            {
                e.Graphics.FillRectangle(background, ClientRectangle);
                e.Graphics.DrawRectangle(edge, 0, 0, Width - 1, Height - 1);
                e.Graphics.FillEllipse(accent, 17, 25, 8, 8);
                Rectangle textBounds = new Rectangle(37, 0, Width - 53, Height);
                using (StringFormat format = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter })
                {
                    e.Graphics.DrawString(message, font, text, textBounds, format);
                }
            }
            base.OnPaint(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { timer.Dispose(); }
            base.Dispose(disposing);
        }
    }

    public static class CaptureToast
    {
        public static void Show(string message)
        {
            using (ToastForm toast = new ToastForm(message)) { toast.ShowDialog(); }
        }
    }
}
