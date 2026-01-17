using System.Drawing;
using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;

namespace PortableWinFormsRecorder;

public sealed class InspectorController : IDisposable
{
    private readonly UIA3Automation _automation = new();
    private readonly System.Windows.Forms.Timer _timer = new();
    private readonly HighlightOverlay _overlay = new();

    public event Action<AutomationElement?>? ElementChanged;

    public bool Enabled
    {
        get => _timer.Enabled;
        set
        {
            if (value)
            {
                if (!_overlay.Visible) _overlay.Show();
                _timer.Start();
            }
            else
            {
                _timer.Stop();
                _overlay.Hide();
                _overlay.SetRect(Rectangle.Empty);
            }
        }
    }

    public InspectorController()
    {
        _timer.Interval = 200;
        _timer.Tick += (_, __) => Tick();
    }

    private void Tick()
    {
        var p = Cursor.Position;
        var el = _automation.FromPoint(p);

        ElementChanged?.Invoke(el);

        try
        {
            if (el != null)
            {
                var r = el.BoundingRectangle;
                if (!r.IsEmpty)
                {
                    var rect = new Rectangle(
                        (int)r.Left - SystemInformation.VirtualScreen.Left,
                        (int)r.Top - SystemInformation.VirtualScreen.Top,
                        (int)r.Width,
                        (int)r.Height
                    );
                    _overlay.SetRect(rect);
                }
            }
        }
        catch { }
    }

    public void Dispose()
    {
        _timer.Stop();
        _timer.Dispose();
        _overlay.Close();
        _overlay.Dispose();
        _automation.Dispose();
    }
}