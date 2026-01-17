using System.Drawing;

namespace PortableWinFormsRecorder;

public sealed class HighlightOverlay : Form
{
    private Rectangle _rect = Rectangle.Empty;

    public HighlightOverlay()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.LimeGreen;
        TransparencyKey = Color.LimeGreen;
        Opacity = 0.6;
        StartPosition = FormStartPosition.Manual;

        Bounds = SystemInformation.VirtualScreen;
    }

    public void SetRect(Rectangle r)
    {
        _rect = r;
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (_rect == Rectangle.Empty) return;

        using var pen = new Pen(Color.Red, 3);
        e.Graphics.DrawRectangle(pen, _rect);
    }
}