using System.Drawing;
using System.Runtime.InteropServices;

namespace PortableWinFormsRecorder;

internal static class WindowCapture
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static (Bitmap Bitmap, Rectangle WindowRect) CaptureForegroundWindowBitmap()
    {
        var hwnd = GetForegroundWindow();
        if (hwnd == IntPtr.Zero)
            throw new Exception("No foreground window.");

        if (!GetWindowRect(hwnd, out var r))
            throw new Exception("GetWindowRect failed.");

        var rect = Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
        if (rect.Width <= 1 || rect.Height <= 1)
            throw new Exception("Foreground window rect is invalid.");

        var bmp = new Bitmap(rect.Width, rect.Height);
        using (var g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(rect.Location, Point.Empty, rect.Size);
        }
        return (bmp, rect);
    }
}
