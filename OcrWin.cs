using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace PortableWinFormsRecorder;

/// <summary>
/// Windows OCR (WinRT) helper (Windows 10/11).
/// </summary>
public static class OcrWin
{

    public readonly record struct OcrBox(string Text, Rectangle Bounds);

public readonly record struct OcrOptions(bool Grayscale, int? Threshold, bool Invert);


    /// <summary>
    /// Captures a screen rectangle and performs OCR. Returns trimmed recognized text.
    /// </summary>
    public static async Task<string> RecognizeScreenRectAsync(Rectangle rect, CancellationToken ct, OcrOptions? options = null)
    {
        using var bmp = new Bitmap(Math.Max(1, rect.Width), Math.Max(1, rect.Height), PixelFormat.Format32bppPArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(rect.Location, Point.Empty, bmp.Size, CopyPixelOperation.SourceCopy);
        }

        if (options != null)
            ApplyPreprocessInPlace(bmp, options.Value);

        var sbmp = ToSoftwareBitmap(bmp);
        ct.ThrowIfCancellationRequested();

        var engine = OcrEngine.TryCreateFromUserProfileLanguages();
        if (engine == null)
            throw new Exception("Windows OCR engine not available on this system.");

        var result = await engine.RecognizeAsync(sbmp);
        return (result?.Text ?? string.Empty).Trim();
    }

    /// <summary>
    /// Performs OCR on an in-memory bitmap and returns word-level boxes.
    /// Bounds are relative to the bitmap's top-left.
    /// </summary>
    public static async Task<List<OcrBox>> RecognizeBitmapBoxesAsync(Bitmap bmp, CancellationToken ct, OcrOptions? options = null)
    {
        if (options != null)
            ApplyPreprocessInPlace(bmp, options.Value);

        var sbmp = ToSoftwareBitmap(bmp);
        ct.ThrowIfCancellationRequested();

        var engine = OcrEngine.TryCreateFromUserProfileLanguages();
        if (engine == null)
            throw new Exception("Windows OCR engine not available on this system.");

        var result = await engine.RecognizeAsync(sbmp);
        var list = new List<OcrBox>();
        if (result == null) return list;

        foreach (var line in result.Lines)
        {
            foreach (var word in line.Words)
            {
                var t = (word.Text ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(t)) continue;
                var r = word.BoundingRect;
                list.Add(new OcrBox(t, new Rectangle((int)r.X, (int)r.Y, (int)r.Width, (int)r.Height)));
            }
        }
        return list;
    }


    private static void ApplyPreprocessInPlace(Bitmap bmp, OcrOptions opt)
    {
        // Very small/fast preprocessing that helps OCR on UI text.
        // - Optional grayscale
        // - Optional binary threshold (0-255)
        // - Optional invert
        var rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
        var data = bmp.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppPArgb);
        try
        {
            unsafe
            {
                byte* ptr = (byte*)data.Scan0.ToPointer();
                int stride = data.Stride;
                for (int y = 0; y < bmp.Height; y++)
                {
                    byte* row = ptr + (y * stride);
                    for (int x = 0; x < bmp.Width; x++)
                    {
                        // BGRA
                        byte b = row[x * 4 + 0];
                        byte g = row[x * 4 + 1];
                        byte r = row[x * 4 + 2];

                        if (opt.Grayscale || opt.Threshold != null)
                        {
                            // Luma approximation
                            int gray = (r * 299 + g * 587 + b * 114) / 1000;
                            r = g = b = (byte)gray;
                        }

                        if (opt.Threshold != null)
                        {
                            int t = Math.Clamp(opt.Threshold.Value, 0, 255);
                            byte v = (byte)((r >= t) ? 255 : 0);
                            r = g = b = v;
                        }

                        if (opt.Invert)
                        {
                            r = (byte)(255 - r);
                            g = (byte)(255 - g);
                            b = (byte)(255 - b);
                        }

                        row[x * 4 + 0] = b;
                        row[x * 4 + 1] = g;
                        row[x * 4 + 2] = r;
                    }
                }
            }
        }
        finally
        {
            bmp.UnlockBits(data);
        }
    }

    private static SoftwareBitmap ToSoftwareBitmap(Bitmap bmp)
    {
        // Convert System.Drawing.Bitmap (BGRA32) -> SoftwareBitmap(Bgra8)
        var rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
        var data = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb);
        try
        {
            int bytes = Math.Abs(data.Stride) * bmp.Height;
            byte[] buffer = new byte[bytes];
            Marshal.Copy(data.Scan0, buffer, 0, bytes);

            // The buffer includes stride padding; SoftwareBitmap expects tightly packed.
            // If stride == width*4, we can use directly; otherwise repack rows.
            int tightStride = bmp.Width * 4;
            if (Math.Abs(data.Stride) == tightStride)
            {
                return SoftwareBitmap.CreateCopyFromBuffer(
                    buffer.AsBuffer(),
                    BitmapPixelFormat.Bgra8,
                    bmp.Width,
                    bmp.Height,
                    BitmapAlphaMode.Premultiplied);
            }

            byte[] tight = new byte[tightStride * bmp.Height];
            int srcStride = Math.Abs(data.Stride);
            for (int y = 0; y < bmp.Height; y++)
            {
                Buffer.BlockCopy(buffer, y * srcStride, tight, y * tightStride, tightStride);
            }

            return SoftwareBitmap.CreateCopyFromBuffer(
                tight.AsBuffer(),
                BitmapPixelFormat.Bgra8,
                bmp.Width,
                bmp.Height,
                BitmapAlphaMode.Premultiplied);
        }
        finally
        {
            bmp.UnlockBits(data);
        }
    }
}
