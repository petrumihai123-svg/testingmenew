using System.Drawing;

namespace PortableWinFormsRecorder;

public sealed class RecordingIndicatorOverlay : Form
{
    private readonly Label _label;

    public RecordingIndicatorOverlay()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;

        BackColor = Color.Black;
        Opacity = 0.65;
        AutoScaleMode = AutoScaleMode.None;

        _label = new Label
        {
            AutoSize = true,
            Text = "RECORDING",
            ForeColor = Color.Red,
            Font = new Font("Segoe UI", 32, FontStyle.Bold),
            Padding = new Padding(20, 10, 20, 10),
            BackColor = Color.Transparent
        };
        Controls.Add(_label);

        // default position (top-center)
        var screen = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 800, 600);
        var w = _label.PreferredWidth;
        var h = _label.PreferredHeight;
        Size = new Size(w, h);
        Location = new Point(screen.Left + (screen.Width - w) / 2, screen.Top + 10);

        // Click-through (doesn't steal focus)
        Enabled = false;
    }

        public void SetText(string text)
    {
        try
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(() => SetText(text));
                return;
            }

            _label.Text = text;
            var w = _label.PreferredWidth;
            var h = _label.PreferredHeight;
            Size = new Size(w, h);

            var screen = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 800, 600);
            Location = new Point(screen.Left + (screen.Width - w) / 2, screen.Top + 10);
        }
        catch { }
    }

    protected override CreateParams CreateParams
    {
        get
        {
            const int WS_EX_TOOLWINDOW = 0x00000080;
            const int WS_EX_TOPMOST = 0x00000008;
            const int WS_EX_TRANSPARENT = 0x00000020;
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_TRANSPARENT;
            return cp;
        }
    }
}
