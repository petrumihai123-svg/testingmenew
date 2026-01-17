using System.Drawing;

namespace PortableWinFormsRecorder;

public sealed class StepFlowView : UserControl
{
    private readonly FlowLayoutPanel _flow = new()
    {
        Dock = DockStyle.Fill,
        AutoScroll = true,
        WrapContents = false,
        FlowDirection = FlowDirection.TopDown
    };

    public StepFlowView()
    {
        Controls.Add(_flow);
    }

    public void ClearAll() => _flow.Controls.Clear();

    public void AddStepThumbnail(string title, string imagePath)
    {
        var panel = new Panel { Width = 520, Height = 120, Margin = new Padding(6) };

        var pic = new PictureBox
        {
            Width = 200,
            Height = 112,
            SizeMode = PictureBoxSizeMode.Zoom,
            BorderStyle = BorderStyle.FixedSingle
        };

        // Load image without locking file
        using (var bmpTemp = new Bitmap(imagePath))
            pic.Image = new Bitmap(bmpTemp);

        var lbl = new Label
        {
            Left = 210,
            Top = 6,
            Width = 300,
            Height = 108,
            AutoEllipsis = true,
            Text = title
        };

        panel.Controls.Add(pic);
        panel.Controls.Add(lbl);

        // Click to open the image
        panel.Cursor = Cursors.Hand;
        panel.Click += (_, __) => TryOpen(imagePath);
        pic.Click += (_, __) => TryOpen(imagePath);
        lbl.Click += (_, __) => TryOpen(imagePath);

        _flow.Controls.Add(panel);
        _flow.ScrollControlIntoView(panel);
    }

    private static void TryOpen(string path)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }
        catch { }
    }
}