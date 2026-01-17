using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DPoint = System.Drawing.Point;
using CvPoint = OpenCvSharp.Point;
using System.Windows.Forms;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace PortableWinFormsRecorder;

public sealed class VideoToFlowWizard : Form
{
    private readonly TextBox _txtVideo = new() { Width = 520 };
    private readonly Button _btnBrowse = new() { Text = "...", Width = 40 };
    private readonly NumericUpDown _numFps = new() { Minimum = 2, Maximum = 30, Value = 10, Width = 60 };
    private readonly NumericUpDown _numMinStopMs = new() { Minimum = 100, Maximum = 2000, Value = 250, Increment = 50, Width = 70 };
    private readonly NumericUpDown _numMergePx = new() { Minimum = 2, Maximum = 100, Value = 20, Width = 60 };
    private readonly NumericUpDown _numMergeMs = new() { Minimum = 100, Maximum = 2000, Value = 600, Increment = 50, Width = 70 };
    private readonly NumericUpDown _numCropW = new() { Minimum = 60, Maximum = 600, Value = 220, Increment = 20, Width = 60 };
    private readonly NumericUpDown _numCropH = new() { Minimum = 60, Maximum = 600, Value = 160, Increment = 20, Width = 60 };
    private readonly Button _btnAnalyze = new() { Text = "Analyze video", Width = 140 };
    private readonly Button _btnAddSteps = new() { Text = "Add to flow", Width = 140, Enabled = false };
    private readonly Button _btnSaveCsv = new() { Text = "Save CSV...", Width = 120, Enabled = false };
    private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
    private readonly PictureBox _preview = new() { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle };
    private readonly Label _status = new() { Dock = DockStyle.Bottom, Height = 22, Text = "" };

    private List<CandidateStep> _candidates = new();
    private string? _videoPath;
    private readonly Dictionary<int, Bitmap> _thumbsByIndex = new();
    private readonly Dictionary<int, string> _imageRelByIndex = new();
    private readonly Dictionary<int, string> _ocrPromptByIndex = new();
    private Bitmap? _firstFrameBmp;

    public sealed record CandidateStep(int TimeMs, int X, int Y, int HoldMs);

    public VideoToFlowWizard()
    {
        Text = "Video → Draft Flow (offline)";
        Width = 980;
        Height = 650;
        StartPosition = FormStartPosition.CenterParent;

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 64, Padding = new Padding(10), AutoScroll = true, WrapContents = false };
        top.Controls.Add(new Label { Text = "Video:", AutoSize = true, Padding = new Padding(0, 8, 0, 0) });
        top.Controls.Add(_txtVideo);
        top.Controls.Add(_btnBrowse);
        top.Controls.Add(new Label { Text = "FPS:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numFps);
        top.Controls.Add(new Label { Text = "Min stop (ms):", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numMinStopMs);
        top.Controls.Add(new Label { Text = "Merge px:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numMergePx);
        top.Controls.Add(new Label { Text = "Merge ms:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numMergeMs);
        top.Controls.Add(new Label { Text = "Crop W:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numCropW);
        top.Controls.Add(new Label { Text = "Crop H:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numCropH);
        top.Controls.Add(_btnAnalyze);
        top.Controls.Add(_btnSaveCsv);
        top.Controls.Add(_btnAddSteps);

        _grid.Columns.Add("Time", "Time (ms)");
        _grid.Columns.Add("X", "X");
        _grid.Columns.Add("Y", "Y");
        _grid.Columns.Add("Hold", "Hold (ms)");

        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 520 };
        split.Panel1.Controls.Add(_grid);
        split.Panel2.Controls.Add(_preview);

        Controls.Add(split);
        Controls.Add(_status);
        Controls.Add(top);

        _btnBrowse.Click += (_, __) => Browse();
        _btnAnalyze.Click += (_, __) => Analyze();
        _btnSaveCsv.Click += async (_, __) => await SaveCsvWithImagesPromptAsync();
        _btnAddSteps.Click += (_, __) => { DialogResult = DialogResult.OK; Close(); };

        _grid.SelectionChanged += (_, __) => ShowPreviewForSelected();
    }

    public IReadOnlyList<CandidateStep> GetCandidates() => _candidates;

    private void Browse()
    {
        using var ofd = new OpenFileDialog
        {
            Filter = "Video files|*.mp4;*.mov;*.mkv;*.avi;*.wmv|All files|*.*",
            Title = "Select a screen recording video"
        };
        if (ofd.ShowDialog(this) == DialogResult.OK)
            _txtVideo.Text = ofd.FileName;
    }

    private void SetStatus(string s)
    {
        _status.Text = s;
        _status.Refresh();
        Application.DoEvents();
    }

    private void Analyze()
    {
        var path = _txtVideo.Text.Trim();
        _videoPath = path;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            MessageBox.Show(this, "Please select a valid video file.", "Video → Flow", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _btnAnalyze.Enabled = false;
        _btnAddSteps.Enabled = false;
        _grid.Rows.Clear();
        _candidates.Clear();
        _firstFrameBmp?.Dispose();
        _firstFrameBmp = null;

        try
        {
            SetStatus("Opening video...");
            using var cap = new VideoCapture(path);
            if (!cap.IsOpened())
                throw new Exception("Could not open video.");

            var fpsWanted = (int)_numFps.Value;
            var srcFps = cap.Fps;
            if (srcFps <= 0) srcFps = fpsWanted;

            var stepFrames = Math.Max(1, (int)Math.Round(srcFps / fpsWanted));
            var minStopMs = (int)_numMinStopMs.Value;
            var mergePx = (int)_numMergePx.Value;
            var mergeMs = (int)_numMergeMs.Value;

            // Improved offline heuristic (fully automatic):
            // 1) Sample frames at FPS.
            // 2) Find "stable" segments (low frame-diff) and pick a keyframe in the middle.
            // 3) Between consecutive keyframes, find the largest changed UI region and use its center as the click point.
            // This is far more reliable than trying to track the cursor in arbitrary recordings.

            using var frame = new Mat();
            int frameIndex = 0;
            int kept = 0;

            // Read first frame for preview
            cap.Read(frame);
            if (frame.Empty()) throw new Exception("Empty video.");
            _firstFrameBmp = BitmapConverter.ToBitmap(frame);

            // Reset to start
            cap.Set(VideoCaptureProperties.PosFrames, 0);

            // Stability detection on downscaled grayscale
            const double stableDiffThresh = 2.0; // avg abs diff on 0..255 scale; tuned for screen recordings
            var stableTimes = new List<int>();
            var keyframes = new List<int>();

            Mat? prevGray = null;
            while (true)
            {
                if (!cap.Read(frame) || frame.Empty())
                    break;

                if (frameIndex % stepFrames != 0)
                {
                    frameIndex++;
                    continue;
                }

                var tMs = (int)cap.PosMsec;

                using var small = new Mat();
                Cv2.Resize(frame, small, new OpenCvSharp.Size(frame.Width / 6, frame.Height / 6));
                using var gray = new Mat();
                Cv2.CvtColor(small, gray, ColorConversionCodes.BGR2GRAY);

                if (prevGray != null)
                {
                    using var diff = new Mat();
                    Cv2.Absdiff(prevGray, gray, diff);
                    var mean = Cv2.Mean(diff).Val0;

                    if (mean <= stableDiffThresh)
                    {
                        stableTimes.Add(tMs);
                    }
                    else
                    {
                        // segment ended; if long enough, commit a keyframe at the median time
                        if (stableTimes.Count >= 2 && (stableTimes[^1] - stableTimes[0]) >= minStopMs)
                        {
                            keyframes.Add(stableTimes[stableTimes.Count / 2]);
                        }
                        stableTimes.Clear();
                        // start a new segment at the current time
                        stableTimes.Add(tMs);
                    }
                }
                else
                {
                    // first sample
                    stableTimes.Add(tMs);
                }

                prevGray?.Dispose();
                prevGray = gray.Clone();

                kept++;
                if (kept % 30 == 0)
                    SetStatus($"Scanning… {tMs / 1000.0:0.0}s");

                frameIndex++;
            }
            prevGray?.Dispose();

            // flush last stable segment
            if (stableTimes.Count >= 2 && (stableTimes[^1] - stableTimes[0]) >= minStopMs)
                keyframes.Add(stableTimes[stableTimes.Count / 2]);

            SetStatus("Building click candidates from UI changes…");

            // Candidate click points are the centers of the largest changed UI region between consecutive keyframes.
            var candidates = new List<CandidateStep>();
            if (keyframes.Count >= 2)
            {
                using var cap2 = new VideoCapture(path);
                using var f1 = new Mat();
                using var f2 = new Mat();

                for (int k = 1; k < keyframes.Count; k++)
                {
                    int tPrev = keyframes[k - 1];
                    int tCur = keyframes[k];

                    cap2.Set(VideoCaptureProperties.PosMsec, tPrev);
                    cap2.Read(f1);
                    cap2.Set(VideoCaptureProperties.PosMsec, tCur);
                    cap2.Read(f2);
                    if (f1.Empty() || f2.Empty()) continue;

                    using var b1 = BitmapConverter.ToBitmap(f1);
                    using var b2 = BitmapConverter.ToBitmap(f2);
                    var p = FindLargestChangeCenter(b1, b2);
                    if (p.HasValue)
                    {
                        // HoldMs is roughly the stable duration around the current keyframe.
                        candidates.Add(new CandidateStep(tCur, p.Value.X, p.Value.Y, Math.Max(0, tCur - tPrev)));
                    }
                }
            }
            else
            {
                // Fallback: if we couldn't find stable segments, keep behavior but warn user.
                SetStatus("No stable segments detected. Try raising FPS or lowering Min stop.");
            }

            // merge close candidates
            candidates = MergeCandidates(candidates, mergeMs, mergePx);

            _candidates = candidates;
            foreach (var c in _candidates)
                _grid.Rows.Add(c.TimeMs, c.X, c.Y, c.HoldMs);

            _btnAddSteps.Enabled = _candidates.Count > 0;
            _btnSaveCsv.Enabled = _candidates.Count > 0;
            BuildThumbnails();
            SetStatus(_candidates.Count == 0
                ? "No candidates found. Try higher FPS or lower Min stop."
                : $"Found {_candidates.Count} candidates. Select rows to preview. Click 'Add to flow' to import.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Video → Flow", MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Error.");
        }
        finally
        {
            _btnAnalyze.Enabled = true;
        }
    }

    private void ShowPreviewForSelected()
    {
        try
        {
            if (_firstFrameBmp == null) return;
            if (_grid.SelectedRows.Count == 0) { _preview.Image = _firstFrameBmp; return; }

            var r = _grid.SelectedRows[0];
            if (r.Cells.Count < 4) return;
            var x = Convert.ToInt32(r.Cells[1].Value);
            var y = Convert.ToInt32(r.Cells[2].Value);

            // Draw crosshair on first frame (good enough for offline preview)
            Bitmap bmp;
            if (_thumbsByIndex.TryGetValue(_grid.SelectedRows[0].Index, out var t))
                bmp = new Bitmap(t);
            else
                bmp = new Bitmap(_firstFrameBmp);
            using var g = Graphics.FromImage(bmp);
            using var pen = new Pen(Color.Lime, 2);
            g.DrawEllipse(pen, x - 12, y - 12, 24, 24);
            g.DrawLine(pen, x - 18, y, x + 18, y);
            g.DrawLine(pen, x, y - 18, x, y + 18);

            _preview.Image?.Dispose();
            _preview.Image = bmp;
        }
        catch { }
    }

    private static double Distance(DPoint a, DPoint b)
    {
        var dx = a.X - b.X;
        var dy = a.Y - b.Y;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    private static DPoint MedianPoint(List<DPoint> pts)
    {
        var xs = pts.Select(p => p.X).OrderBy(x => x).ToArray();
        var ys = pts.Select(p => p.Y).OrderBy(y => y).ToArray();
        return new DPoint(xs[xs.Length / 2], ys[ys.Length / 2]);
    }

    private static List<CandidateStep> MergeCandidates(List<CandidateStep> input, int mergeMs, int mergePx)
    {
        if (input.Count <= 1) return input;
        input = input.OrderBy(c => c.TimeMs).ToList();
        var outp = new List<CandidateStep> { input[0] };

        foreach (var c in input.Skip(1))
        {
            var last = outp[^1];
            var dt = c.TimeMs - last.TimeMs;
            var dx = Math.Abs(c.X - last.X);
            var dy = Math.Abs(c.Y - last.Y);
            if (dt <= mergeMs && dx <= mergePx && dy <= mergePx)
            {
                // keep the longer hold and newer time (roughly)
                var merged = last with
                {
                    TimeMs = c.TimeMs,
                    X = (last.X + c.X) / 2,
                    Y = (last.Y + c.Y) / 2,
                    HoldMs = Math.Max(last.HoldMs, c.HoldMs)
                };
                outp[^1] = merged;
            }
            else
            {
                outp.Add(c);
            }
        }
        return outp;
    }


    private async Task SaveCsvWithImagesPromptAsync()
    {
        try
        {
            using var sfd = new SaveFileDialog
            {
                Filter = "CSV files|*.csv|All files|*.*",
                FileName = "video_flow.csv",
                Title = "Save draft flow CSV (with extracted images)"
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;
            _btnSaveCsv.Enabled = false;
            _btnAnalyze.Enabled = false;
            _btnAddSteps.Enabled = false;
            SetStatus("Saving CSV + extracting images...");

            // Extraction (and OCR) can take time; run off the UI thread to avoid freezing.
            var outPath = sfd.FileName;
            await Task.Run(() => SaveCsvWithImages(outPath));

            SetStatus("Saved.");
            MessageBox.Show(this, "Saved CSV + images successfully.", "Video → Flow", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Video → Flow", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _btnAnalyze.Enabled = true;
            _btnSaveCsv.Enabled = _candidates.Count > 0;
            _btnAddSteps.Enabled = _candidates.Count > 0;
        }
    }

    public void SaveCsvWithImages(string csvPath)
    {
        if (_candidates.Count == 0) throw new Exception("No candidates to save.");
        if (string.IsNullOrWhiteSpace(csvPath)) throw new Exception("Invalid CSV path.");

        var assetsDir = Assets.ResolveAssetsDir(csvPath);
        Directory.CreateDirectory(assetsDir);

        var script = BuildScriptWithImages(assetsDir);
        ScriptCsv.Save(script, csvPath);
    }

    public Script BuildScriptWithImages(string assetsDir)
    {
        if (_videoPath == null) throw new Exception("No video loaded.");
        if (_candidates.Count == 0) return new Script();

        Directory.CreateDirectory(assetsDir);

        // Extract template images for each candidate at its timestamp.
        ExtractTemplatesIntoAssets(assetsDir);

        var script = new Script();
        // Conservative: delay + click steps. Click steps use Target.Image (image search fallback).
        var steps = new List<Step>();

        steps.Add(new Step { Action = "Delay", Value = "300" });

        int lastT = _candidates[0].TimeMs;
        for (int i = 0; i < _candidates.Count; i++)
        {
            var c = _candidates[i];
            var delay = Math.Max(0, c.TimeMs - lastT);
            if (delay > 50)
                steps.Add(new Step { Action = "Delay", Value = delay.ToString() });

            var imgRel = _imageRelByIndex.TryGetValue(i, out var rel) ? rel : null;

            var step = new Step
            {
                Action = "Click",
                // Always include absolute click coords in Value as a last-resort fallback.
                // Runner understands Value="x,y".
                Value = $"{c.X},{c.Y}",
                Note = "video-draft",
                Target = new Target
                {
                    Image = imgRel,
                    // Best-effort OCR hint from the template crop.
                    // Runner can use this for SmartFind fallback (OCR-on-active-window).
                    Name = _ocrPromptByIndex.TryGetValue(i, out var prompt) ? prompt : null
                }
            };

            steps.Add(step);
            lastT = c.TimeMs;
        }

        script.Steps = steps;
        return script;
    }

    private void BuildThumbnails()
    {
        try
        {
            _thumbsByIndex.Clear();
            if (_videoPath == null || !File.Exists(_videoPath)) return;

            using var cap = new VideoCapture(_videoPath);
            if (!cap.IsOpened()) return;

            using var frame = new Mat();

            var cropW = (int)_numCropW.Value;
            var cropH = (int)_numCropH.Value;

            for (int i = 0; i < _candidates.Count; i++)
            {
                cap.Set(VideoCaptureProperties.PosMsec, _candidates[i].TimeMs);
                cap.Read(frame);
                if (frame.Empty()) continue;

                using var bmpFull = BitmapConverter.ToBitmap(frame);
                var thumb = CropAround(bmpFull, _candidates[i].X, _candidates[i].Y, cropW, cropH);
                _thumbsByIndex[i] = thumb;
            }
        }
        catch { }
    }

    private void ExtractTemplatesIntoAssets(string assetsDir)
    {
        _imageRelByIndex.Clear();
        _ocrPromptByIndex.Clear();
        if (_videoPath == null || !File.Exists(_videoPath)) return;

        using var cap = new VideoCapture(_videoPath);
        if (!cap.IsOpened()) return;

        using var frame = new Mat();

        var cropW = (int)_numCropW.Value;
        var cropH = (int)_numCropH.Value;

        for (int i = 0; i < _candidates.Count; i++)
        {
            cap.Set(VideoCaptureProperties.PosMsec, _candidates[i].TimeMs);
            cap.Read(frame);
            if (frame.Empty()) continue;

            using var bmpFull = BitmapConverter.ToBitmap(frame);
            using var crop = CropAround(bmpFull, _candidates[i].X, _candidates[i].Y, cropW, cropH);

            var rel = $"video_step_{i + 1:000}.png";
            var full = Path.Combine(assetsDir, rel);
            crop.Save(full, System.Drawing.Imaging.ImageFormat.Png);
            _imageRelByIndex[i] = rel;

            // Try extracting a text hint from the crop using Windows OCR.
            // This enables runner SmartFind fallback (OCR-on-active-window) even when template matching fails.
            try
            {
                var prompt = TryOcrPromptFromCrop(crop);
                if (!string.IsNullOrWhiteSpace(prompt))
                    _ocrPromptByIndex[i] = prompt;
            }
            catch
            {
                // ignore OCR failures; this wizard should still work without OCR.
            }
        }
    }

    private static string? TryOcrPromptFromCrop(Bitmap crop)
    {
        // Best-effort: OCR the crop and return the most "button-like" line.
        // Uses word boxes only (no confidence exposed by WinRT OCR).
        using var bmp = new Bitmap(crop);
        // Avoid deadlocks by not capturing the WinForms synchronization context.
        var boxes = OcrWin.RecognizeBitmapBoxesAsync(bmp, CancellationToken.None, new OcrWin.OcrOptions(Grayscale: true, Threshold: null, Invert: false))
            .ConfigureAwait(false).GetAwaiter().GetResult();
        if (boxes.Count == 0) return null;

        // Group words into lines by Y proximity.
        var words = boxes
            .Select(b => (Text: b.Text.Trim(), Bounds: b.Bounds))
            .Where(w => w.Text.Length > 0)
            .ToList();
        if (words.Count == 0) return null;

        const int lineYThresh = 10;
        var lines = new List<List<(string Text, Rectangle Bounds)>>();
        foreach (var w in words.OrderBy(w => w.Bounds.Top).ThenBy(w => w.Bounds.Left))
        {
            var placed = false;
            foreach (var line in lines)
            {
                var y = line[0].Bounds.Top;
                if (Math.Abs(w.Bounds.Top - y) <= lineYThresh)
                {
                    line.Add((w.Text, w.Bounds));
                    placed = true;
                    break;
                }
            }
            if (!placed)
                lines.Add(new List<(string Text, Rectangle Bounds)> { (w.Text, w.Bounds) });
        }

        string Clean(string s)
        {
            s = s.Trim();
            // remove very short noise tokens
            s = new string(s.Where(ch => char.IsLetterOrDigit(ch) || char.IsWhiteSpace(ch) || ch == '-' || ch == '_' || ch == '.').ToArray());
            return s.Trim();
        }

        // Score lines: prefer moderate length, mostly letters, not giant paragraphs.
        (string Text, double Score) best = ("", double.MinValue);
        foreach (var line in lines)
        {
            var text = string.Join(" ", line.OrderBy(w => w.Bounds.Left).Select(w => w.Text));
            text = Clean(text);
            if (string.IsNullOrWhiteSpace(text)) continue;
            if (text.Length < 2) continue;
            if (text.Length > 40) continue;

            int letters = text.Count(char.IsLetter);
            int digits = text.Count(char.IsDigit);
            int spaces = text.Count(char.IsWhiteSpace);
            double letterRatio = letters / (double)Math.Max(1, text.Length - spaces);

            double score = 0;
            score += letterRatio * 2.0;
            score += Math.Min(1.0, text.Length / 12.0);
            score -= digits * 0.05;
            // Prefer TitleCase-ish and common UI tokens.
            if (Regex.IsMatch(text, "^(OK|Ok|Cancel|Yes|No|Next|Back|Finish|Save|Open|Close|Run|Start)$")) score += 1.0;
            if (char.IsUpper(text[0])) score += 0.2;

            if (score > best.Score)
                best = (text, score);
        }

        return string.IsNullOrWhiteSpace(best.Text) ? null : best.Text;
    }

    private static Bitmap CropAround(Bitmap bmp, int x, int y, int w, int h)
    {
        w = Math.Max(20, Math.Min(w, bmp.Width));
        h = Math.Max(20, Math.Min(h, bmp.Height));

        int left = x - (w / 2);
        int top = y - (h / 2);

        if (left < 0) left = 0;
        if (top < 0) top = 0;
        if (left + w > bmp.Width) left = bmp.Width - w;
        if (top + h > bmp.Height) top = bmp.Height - h;

        var rect = new Rectangle(left, top, w, h);
        var outBmp = new Bitmap(rect.Width, rect.Height);
        using var g = Graphics.FromImage(outBmp);
        g.DrawImage(bmp, new Rectangle(0, 0, rect.Width, rect.Height), rect, GraphicsUnit.Pixel);
        return outBmp;
    }

    private static DPoint? FindLargestChangeCenter(Bitmap before, Bitmap after)
    {
        // Find the biggest "changed" region between two keyframes.
        // This tends to correspond to the UI element that was interacted with.
        // Returns a point in full-resolution screen coordinates.

        try
        {
            using var m1 = BitmapConverter.ToMat(before);
            using var m2 = BitmapConverter.ToMat(after);

            // Work at half-res for speed/robustness.
            using var s1 = new Mat();
            using var s2 = new Mat();
            Cv2.Resize(m1, s1, new OpenCvSharp.Size(m1.Width / 2, m1.Height / 2));
            Cv2.Resize(m2, s2, new OpenCvSharp.Size(m2.Width / 2, m2.Height / 2));

            using var d = new Mat();
            Cv2.Absdiff(s1, s2, d);

            using var g = new Mat();
            Cv2.CvtColor(d, g, ColorConversionCodes.BGR2GRAY);

            // Threshold: Otsu works well across videos.
            using var bw = new Mat();
            Cv2.Threshold(g, bw, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);

            // Remove tiny noise and join nearby pixels.
            using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(7, 7));
            Cv2.MorphologyEx(bw, bw, MorphTypes.Open, kernel);
            Cv2.MorphologyEx(bw, bw, MorphTypes.Close, kernel);

            Cv2.FindContours(bw, out var contours, out _, RetrievalModes.External, ContourApproximationModes.ApproxSimple);
            if (contours.Length == 0) return null;

            var frameArea = (double)bw.Width * bw.Height;
            OpenCvSharp.Rect best = default;
            double bestArea = 0;

            foreach (var c in contours)
            {
                var r = Cv2.BoundingRect(c);
                var a = (double)r.Width * r.Height;

                // Ignore very small noise
                if (a < frameArea * 0.001) continue;
                // Ignore "everything changed" (window switch) – too ambiguous
                if (a > frameArea * 0.65) continue;

                if (a > bestArea)
                {
                    bestArea = a;
                    best = r;
                }
            }

            if (bestArea <= 0) return null;

            // Center at half-res, then scale back up.
            var cx = (best.Left + best.Width / 2) * 2;
            var cy = (best.Top + best.Height / 2) * 2;
            cx = Math.Clamp(cx, 0, before.Width - 1);
            cy = Math.Clamp(cy, 0, before.Height - 1);
            return new DPoint(cx, cy);
        }
        catch
        {
            return null;
        }
    }

}