using System.Drawing;
using System.Windows.Forms;

namespace PortableWinFormsRecorder;

/// <summary>
/// Fullscreen translucent overlay that lets the user drag-select a screen region.
/// Returns a Rectangle in SCREEN coordinates (virtual screen).
/// </summary>
public sealed class RegionPickerOverlay : Form
{
    private Point _start;
    private Point _current;
    private bool _dragging;

    public Rectangle SelectedRect { get; private set; }

    public RegionPickerOverlay()
    {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        TopMost = true;
        ShowInTaskbar = false;
        DoubleBuffered = true;
        Cursor = Cursors.Cross;

        var vs = SystemInformation.VirtualScreen;
        Bounds = vs;

        BackColor = Color.Black;
        Opacity = 0.20;
        KeyPreview = true;

        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        };

        MouseDown += (_, e) =>
        {
            if (e.Button != MouseButtons.Left) return;
            _dragging = true;
            _start = PointToScreen(e.Location);
            _current = _start;
            Invalidate();
        };

        MouseMove += (_, e) =>
        {
            if (!_dragging) return;
            _current = PointToScreen(e.Location);
            Invalidate();
        };

        MouseUp += (_, e) =>
        {
            if (!_dragging || e.Button != MouseButtons.Left) return;
            _dragging = false;
            _current = PointToScreen(e.Location);

            SelectedRect = NormalizeRect(_start, _current);
            if (SelectedRect.Width < 2 || SelectedRect.Height < 2)
            {
                DialogResult = DialogResult.Cancel;
            }
            else
            {
                DialogResult = DialogResult.OK;
            }
            Close();
        };
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        if (!_dragging) return;

        var r = NormalizeRect(_start, _current);
        if (r.Width <= 0 || r.Height <= 0) return;

        // Convert screen rect -> overlay client coords
        r.Offset(-Bounds.X, -Bounds.Y);

        using var pen = new Pen(Color.Lime, 2);
        using var brush = new SolidBrush(Color.FromArgb(60, Color.Lime));
        e.Graphics.FillRectangle(brush, r);
        e.Graphics.DrawRectangle(pen, r);
    }

    private static Rectangle NormalizeRect(Point a, Point b)
    {
        int x1 = Math.Min(a.X, b.X);
        int y1 = Math.Min(a.Y, b.Y);
        int x2 = Math.Max(a.X, b.X);
        int y2 = Math.Max(a.Y, b.Y);
        return new Rectangle(x1, y1, x2 - x1, y2 - y1);
    }
}
