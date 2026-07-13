using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ToolRackProbe
{
    public sealed class LatencyForm : Form
    {
        private bool recorded;

        public LatencyForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-30000, -30000);
            ClientSize = new Size(40, 40);
            BackColor = Color.White;
        }

        protected override void OnShown(EventArgs eventArgs)
        {
            base.OnShown(eventArgs);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs eventArgs)
        {
            base.OnPaint(eventArgs);
            if (recorded)
            {
                return;
            }
            recorded = true;
            string marker = Environment.GetEnvironmentVariable("TOOLRACK_LATENCY_MARKER");
            if (!String.IsNullOrEmpty(marker))
            {
                string value = DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
                File.WriteAllText(marker, value, new UTF8Encoding(false));
            }
            BeginInvoke(new MethodInvoker(Close));
        }
    }

    public static class LatencyPalette
    {
        public static int Run()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new LatencyForm());
            return 0;
        }
    }
}
