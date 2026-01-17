using System.Drawing;
using System.Drawing.Imaging;

namespace PortableWinFormsRecorder;

public static class ImageSearch
{
    public sealed record Match(Point Location, Size Size, double Score);

    public static Bitmap CaptureVirtualScreen()
    {
        var bounds = SystemInformation.VirtualScreen;
        var bmp = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format24bppRgb);
        using var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size);
        return bmp;
    }

    public static Match? FindOnScreen(Bitmap screen, Bitmap template, double maxNormalizedSad = 0.08)
    {
        if (template.Width < 5 || template.Height < 5) return null;
        if (template.Width > screen.Width || template.Height > screen.Height) return null;

        var s = ToGray(screen);
        var t = ToGray(template);

        int sw = screen.Width, sh = screen.Height;
        int tw = template.Width, th = template.Height;

        long bestSad = long.MaxValue;
        int bestX = -1, bestY = -1;

        long maxSad = (long)tw * th * 255;

        for (int y = 0; y <= sh - th; y++)
        {
            for (int x = 0; x <= sw - tw; x++)
            {
                long sad = 0;

                for (int ty = 0; ty < th; ty++)
                {
                    int si = (y + ty) * sw + x;
                    int ti = ty * tw;

                    for (int tx = 0; tx < tw; tx++)
                    {
                        int diff = s[si + tx] - t[ti + tx];
                        if (diff < 0) diff = -diff;
                        sad += diff;
                        if (sad >= bestSad) goto NextPos;
                    }
                }

                if (sad < bestSad)
                {
                    bestSad = sad;
                    bestX = x;
                    bestY = y;
                }

            NextPos:
                ;
            }
        }

        if (bestX < 0) return null;

        double norm = (double)bestSad / maxSad;
        if (norm > maxNormalizedSad) return null;

        return new Match(new Point(bestX + SystemInformation.VirtualScreen.Left, bestY + SystemInformation.VirtualScreen.Top),
                         new Size(tw, th),
                         norm);
    }

    /// <summary>
    /// Multi-scale template matching to tolerate DPI/theme scaling differences.
    /// Tries a small range of scales around 1.0 and returns the best match under the given threshold.
    /// </summary>
    public static Match? FindOnScreenMultiScale(Bitmap screen, Bitmap template, double maxNormalizedSad = 0.08)
    {
        // Keep the scale set small; this is brute-force SAD.
        var scales = new double[] { 0.85, 0.90, 0.95, 1.00, 1.05, 1.10, 1.15 };
        Match? best = null;
        foreach (var s in scales)
        {
            using var scaled = s == 1.0 ? null : Scale(template, s);
            var temp = scaled ?? template;
            var m = FindOnScreen(screen, temp, maxNormalizedSad);
            if (m == null) continue;
            if (best == null || m.Score < best.Score)
                best = m;
        }
        return best;
    }

    private static Bitmap Scale(Bitmap src, double scale)
    {
        var w = Math.Max(1, (int)Math.Round(src.Width * scale));
        var h = Math.Max(1, (int)Math.Round(src.Height * scale));
        var dst = new Bitmap(w, h);
        using var g = Graphics.FromImage(dst);
        // Nearest-neighbor-ish: keep it simple and fast; UI templates usually have sharp edges.
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
        g.DrawImage(src, new Rectangle(0, 0, w, h));
        return dst;
    }

    private static byte[] ToGray(Bitmap bmp)
    {
        using var clone = new Bitmap(bmp.Width, bmp.Height, PixelFormat.Format24bppRgb);
        using (var g = Graphics.FromImage(clone))
            g.DrawImageUnscaled(bmp, 0, 0);

        var rect = new Rectangle(0, 0, clone.Width, clone.Height);
        var data = clone.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

        try
        {
            int w = clone.Width, h = clone.Height;
            byte[] gray = new byte[w * h];

            unsafe
            {
                byte* ptr = (byte*)data.Scan0;
                int stride = data.Stride;

                for (int y = 0; y < h; y++)
                {
                    byte* row = ptr + y * stride;
                    int oi = y * w;
                    for (int x = 0; x < w; x++)
                    {
                        byte b = row[x * 3 + 0];
                        byte g2 = row[x * 3 + 1];
                        byte r = row[x * 3 + 2];
                        gray[oi + x] = (byte)((r * 77 + g2 * 150 + b * 29) >> 8);
                    }
                }
            }

            return gray;
        }
        finally
        {
            clone.UnlockBits(data);
        }
    }
}