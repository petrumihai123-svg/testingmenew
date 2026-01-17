using System.Diagnostics;
using System.Drawing;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PortableWinFormsRecorder;

public sealed class RunnerOptions
{
    public string AssetsDir { get; set; } = ".";
    public Action<string>? Log { get; set; }

    // Stop hotkey (shared with Recorder)
    // Defaults match the UI: Ctrl+Shift+S
    public bool StopHotkeyCtrl { get; set; } = true;
    public bool StopHotkeyShift { get; set; } = true;
    public bool StopHotkeyAlt { get; set; } = false;
    public char StopHotkeyKey { get; set; } = 'S';

    // Run controls
    public int StartIndex { get; set; } = 0;                // run from this step index
    public int RetryCount { get; set; } = 0;                // retries per step on failure
    public int RetryDelayMs { get; set; } = 250;            // delay between retries

    // UI progress
    public Action<int>? OnStepStart { get; set; }           // called with step index before execution

    // Image matching
    public double ImageMaxNormalizedSad { get; set; } = 0.08;

    // Offline "AI" fallback: use Windows OCR bounding boxes + fuzzy matching
    // when UIAutomation fails (timeouts etc.).
    public bool UseSmartFindFallback { get; set; } = true;
}

public static class Runner
{
    public static void RunOnce(Script script, Dictionary<string, string> data, RunnerOptions? opts = null)
    {
        opts ??= new RunnerOptions();
        void Log(string s) => opts.Log?.Invoke(s);

        // Diagnostics bundle for this run (screenshots, match overlays, etc.).
        // Created eagerly so any failure can write artifacts that explain what happened.
        var runDir = CreateRunDiagnosticsDir(opts.AssetsDir);

        using var stopCts = new CancellationTokenSource();
        using var stopPoller = StartStopHotkeyPoller(opts, () =>
        {
            if (stopCts.IsCancellationRequested) return;
            stopCts.Cancel();
            Log("[RUN] stop requested (hotkey)");
        });

        void ThrowIfStopped()
        {
            if (stopCts.IsCancellationRequested)
                throw new OperationCanceledException("Run stopped by hotkey.");
        }

        using var automation = new UIA3Automation();

        AutomationElement? scopeRoot = null;
        Window? window = null;

        // Track last click position (screen coordinates). Used for relative OCR regions: rel:dx,dy,w,h
        Point? lastClick = null;

        if (!string.IsNullOrWhiteSpace(script.App.ProcessName))
        {
            var proc = Process.GetProcessesByName(script.App.ProcessName).FirstOrDefault()
                       ?? throw new Exception($"Process not found: {script.App.ProcessName}");

            using var app = FlaUI.Core.Application.Attach(proc);

            window = Wait(() => app.GetMainWindow(automation), TimeSpan.FromSeconds(10))
                     ?? throw new Exception("Main window not found");

            window.Focus();
            scopeRoot = window;
        }
        else
        {
            // Desktop mode: operate from global desktop root
            scopeRoot = automation.GetDesktop();
        }

        var steps = script.Steps ?? new List<Step>();
        for (int i = Math.Max(0, opts.StartIndex); i < steps.Count; i++)
        {
            ThrowIfStopped();
            var step = steps[i];

            opts.OnStepStart?.Invoke(i);

            if (step.DelayMs is > 0)
                SleepWithStop(step.DelayMs.Value, stopCts.Token);

            int attempt = 0;
            while (true)
            {
                ThrowIfStopped();
                try
                {
                    AutomationElement? el = null;
                    Exception? findEx = null;
                    try
                    {
                        el = Find(scopeRoot!, step.Target);
                    }
                    catch (System.Runtime.InteropServices.COMException ce)
                    {
                        // UIAutomation can intermittently time out on heavy UI trees.
                        // Treat it as "not found" and allow fallbacks.
                        findEx = ce;
                        el = null;
                    }
                    catch (Exception ex0)
                    {
                        findEx = ex0;
                        el = null;
                    }

                    // Offline smart fallback (OCR bounding boxes + fuzzy matching) only uses active window content.
                    SmartFindMatch? smartMatch = null;
                    if (el == null && opts.UseSmartFindFallback)
                    {
                        // Prefer target.Name; fall back to step.Value prompt for SmartFindClick.
                        var prompt = step.Target?.Name;
                        if (string.IsNullOrWhiteSpace(prompt) && step.Action == "SmartFindClick")
                            prompt = ParseKeyVals(step.Value ?? "").GetValueOrDefault("prompt");

                        if (!string.IsNullOrWhiteSpace(prompt))
                            smartMatch = TrySmartFindOnActiveWindow(prompt!, stopCts.Token);
                    }

                    ImageSearch.Match? imgMatch = null;
                    var imgRel = step.Target?.Image;
                    if (el == null && !string.IsNullOrWhiteSpace(imgRel))
                    {
                        var imgPath = Assets.Combine(opts.AssetsDir, imgRel);
                        if (File.Exists(imgPath))
                        {
                            using var temp = new Bitmap(imgPath);
                            using var screen = ImageSearch.CaptureVirtualScreen();
                            imgMatch = ImageSearch.FindOnScreenMultiScale(screen, temp, opts.ImageMaxNormalizedSad);
                            if (imgMatch != null)
                                Log($"[RUN] image fallback match: x={(imgMatch.Location.X + imgMatch.Size.Width / 2)}, y={(imgMatch.Location.Y + imgMatch.Size.Height / 2)}, sad={imgMatch.Score:0.000}");
                        }
                    }

                    // Coordinate fallback from Value, format: "x,y" (screen coordinates)
                    // This is especially useful for video-draft flows where UIA/Template matching can fail.
                    Point? valuePoint = null;
                    if (el == null && smartMatch == null && imgMatch == null)
                        valuePoint = TryParsePoint(step.Value);

                    switch (step.Action)
                    {
                    case "Delay":
                        {
                            var ms = 0;
                            int.TryParse((step.Value ?? "0").Trim(), out ms);
                            if (ms < 0) ms = 0;
                            SleepWithStop(ms, stopCts.Token);
                            break;
                        }


                        case "Click":
                            if (el != null)
                                lastClick = Click(el);
                            else if (smartMatch != null)
                            {
                                Mouse.Click(smartMatch.Value.ClickPoint);
                                lastClick = smartMatch.Value.ClickPoint;
                                Log($"[RUN] smart fallback click: '{smartMatch.Value.MatchedText}' score={smartMatch.Value.Score:0.00} at {smartMatch.Value.ClickPoint.X},{smartMatch.Value.ClickPoint.Y}");
                            }
                            else if (imgMatch != null)
                            {
                                var p = new Point(imgMatch.Location.X + imgMatch.Size.Width / 2, imgMatch.Location.Y + imgMatch.Size.Height / 2);
                                Mouse.Click(p);
                                lastClick = p;
                            }
                            else if (valuePoint != null)
                            {
                                Mouse.Click(valuePoint.Value);
                                lastClick = valuePoint.Value;
                                Log($"[RUN] coord fallback click: {valuePoint.Value.X},{valuePoint.Value.Y}");
                            }
                            else
                            {
                                WriteStepDiagnostics(runDir, i, step, imgRel, findEx);
                                throw new Exception(findEx != null
                                    ? $"Click target not found (UIA failed: {findEx.Message}). Diagnostics: {runDir}"
                                    : $"Click target not found (no UIA element and no image match). Diagnostics: {runDir}" );
                            }
                            break;

                        case "SetText":
                            var value = Templating.Apply(step.Value ?? "", data);
                            if (el != null)
                                SetText(el, value);
                            else if (smartMatch != null)
                            {
                                Mouse.Click(smartMatch.Value.ClickPoint);
                                lastClick = smartMatch.Value.ClickPoint;
                                Keyboard.Type(value);
                                Log($"[RUN] smart fallback settext: '{smartMatch.Value.MatchedText}' score={smartMatch.Value.Score:0.00} at {smartMatch.Value.ClickPoint.X},{smartMatch.Value.ClickPoint.Y}");
                            }
                            else if (imgMatch != null)
                            {
                                var p = new Point(imgMatch.Location.X + imgMatch.Size.Width / 2, imgMatch.Location.Y + imgMatch.Size.Height / 2);
                                Mouse.Click(p);
                                lastClick = p;
                                Keyboard.Type(value);
                            }
                            else if (valuePoint != null)
                            {
                                Mouse.Click(valuePoint.Value);
                                lastClick = valuePoint.Value;
                                Keyboard.Type(value);
                                Log($"[RUN] coord fallback settext: {valuePoint.Value.X},{valuePoint.Value.Y}");
                            }
                            else
                            {
                                WriteStepDiagnostics(runDir, i, step, imgRel, findEx);
                                throw new Exception(findEx != null
                                    ? $"SetText target not found (UIA failed: {findEx.Message}). Diagnostics: {runDir}"
                                    : $"SetText target not found (no UIA element and no image match). Diagnostics: {runDir}" );
                            }
                            break;

                        // Type raw text into the currently focused control/window.
                        // Useful when recording keyboard input in apps where the editable UIA element
                        // cannot be reliably identified (e.g., modern Notepad/custom editors).
                        case "TypeText":
                            var t = Templating.Apply(step.Value ?? "", data);
                            Keyboard.Type(t);
                            Log($"[RUN] TypeText: {t}");
                            break;

                        case "AssertTextEquals":
                            var expected = Templating.Apply(step.Value ?? "", data);
                            var actual = GetText(el);
                            if (!string.Equals(expected, actual, StringComparison.Ordinal))
                                throw new Exception($"Assertion failed. Expected \"{expected}\" but got \"{actual}\"");
                            break;

                        case "AssertExists":
                            if (el == null && imgMatch == null)
                                throw new Exception("Assertion failed: target does not exist (no UIA element and no image match).");
                            break;

                        // Offline smart find + click using OCR bounding boxes.
                        // Step.Value format: prompt=...|minScore=0.6|fuzzy=1
                        case "SmartFindClick":
                        {
                            var kv = ParseKeyVals(Templating.Apply(step.Value ?? "", data));
                            var prompt = kv.GetValueOrDefault("prompt") ?? step.Target?.Name ?? "";
                            if (string.IsNullOrWhiteSpace(prompt))
                                throw new Exception("SmartFindClick requires prompt=... (or Target.Name).");

                            var minScore = ParseDouble(kv.GetValueOrDefault("minScore"), 0.60);
                            var fuzzy = ParseBool(kv.GetValueOrDefault("fuzzy"), true);

                            var m = TrySmartFindOnActiveWindow(prompt, stopCts.Token, minScore: minScore, allowFuzzy: fuzzy);
                            if (m == null)
                                throw new Exception($"SmartFindClick: could not find '{prompt}' on active window.");

                            Mouse.Click(m.Value.ClickPoint);
                            lastClick = m.Value.ClickPoint;
                            Log($"[RUN] SmartFindClick '{m.Value.MatchedText}' score={m.Value.Score:0.00} at {m.Value.ClickPoint.X},{m.Value.ClickPoint.Y}");
                            break;
                        }

                        // Offline smart extract: OCR region then extract digits.
                        // Step.Value format:
                        //   region=x,y,w,h OR region=rel:dx,dy,w,h | save=Var | want=11digits/8digits/7digits | cleanup=digits
                        case "SmartExtract":
                        {
                            var kv = ParseKeyVals(Templating.Apply(step.Value ?? "", data));
                            var regionSpec = kv.GetValueOrDefault("region") ?? (step.Value ?? "");
                            if (regionSpec.StartsWith("region=", StringComparison.OrdinalIgnoreCase))
                                regionSpec = regionSpec.Substring("region=".Length);
                            var save = kv.GetValueOrDefault("save") ?? "Text";
                            var want = (kv.GetValueOrDefault("want") ?? "").Trim().ToLowerInvariant();
                            var cleanup = (kv.GetValueOrDefault("cleanup") ?? "").Trim().ToLowerInvariant();
                            var parsed = ParseOcrSpec(regionSpec, lastClick);

                            var raw = OcrWin.RecognizeScreenRectAsync(parsed.Rect, stopCts.Token, parsed.Options)
                                .GetAwaiter().GetResult();
                            var s = raw ?? "";
                            if (cleanup == "digits")
                                s = Regex.Replace(s, "\\D+", "");
                            else
                                s = s.Trim();

                            int? n = want switch
                            {
                                "11digits" => 11,
                                "8digits" => 8,
                                "7digits" => 7,
                                _ => null
                            };

                            if (n != null)
                            {
                                var mm = Regex.Match(s, $"\\d{{{n.Value}}}");
                                s = mm.Success ? mm.Value : "";
                            }

                            data[save] = s;
                            Log($"[RUN] SmartExtract -> {save}='{TruncateForLog(s)}'");
                            break;
                        }

                        // Windows OCR: capture a screen rectangle and save recognized text into a variable.
                        // Step.Value format:
                        //   x,y,w,h|save=VarName|regex=...   (regex is optional)
                        // Example:
                        //   100,200,260,60|save=Code|regex=\b\d{6}\b
                        
                        case "OcrRead":
                        {
                            var parsed = ParseOcrSpec(Templating.Apply(step.Value ?? "", data), lastClick);

                            string text = "";
                            int tries = Math.Max(1, parsed.Retries);
                            for (int attempt = 0; attempt < tries; attempt++)
                            {
                                ThrowIfStopped();
                                try
                                {
                                    text = OcrWin.RecognizeScreenRectAsync(parsed.Rect, stopCts.Token, parsed.Options)
                                        .GetAwaiter().GetResult();
                                }
                                catch
                                {
                                    if (attempt == tries - 1) throw;
                                }

                                if (!string.IsNullOrWhiteSpace(parsed.Regex))
                                {
                                    var mm = Regex.Match(text ?? "", parsed.Regex, RegexOptions.Multiline);
                                    text = mm.Success ? mm.Value : "";
                                }

                                if (!string.IsNullOrWhiteSpace(text) || attempt == tries - 1)
                                    break;

                                SleepWithStop(Math.Max(0, parsed.RetryDelayMs), stopCts.Token);
                            }

                            data[parsed.SaveVar] = (text ?? "").Trim();
                            Log($"[RUN] OCR -> {parsed.SaveVar}='{TruncateForLog(text ?? "")}' rect={parsed.Rect.X},{parsed.Rect.Y},{parsed.Rect.Width},{parsed.Rect.Height}");
                            break;
                        }
// Paste a variable/template into the currently focused field.
                        // Step.Value can be a template, e.g. {{Code}}
                        case "PasteVar":
                        {
                            var text = Templating.Apply(step.Value ?? "", data);
                            PasteText(text);
                            break;
                        }



                        default:
                            Log($"Unknown action: {step.Action}");
                            break;
                    }

                    // success
                    break;
                }
                catch (Exception ex)
                {
                    ThrowIfStopped();
                    if (attempt >= opts.RetryCount)
                        throw;

                    attempt++;
                    Log($"[RUN] Step {i} failed (attempt {attempt}/{opts.RetryCount}). Retrying... {ex.Message}");
                    SleepWithStop(Math.Max(0, opts.RetryDelayMs), stopCts.Token);
                }
            }
        }
    }

    private readonly record struct OcrSpec(Rectangle Rect, string SaveVar, string? Regex, int Retries, int RetryDelayMs, OcrWin.OcrOptions Options);

    
    private static OcrSpec ParseOcrSpec(string spec, Point? lastClick)
    {
        // Expected:
        //   x,y,w,h|save=VarName|regex=...|retries=3|retryDelay=150|pre=gray|thresh=180|invert=1
        // Relative to last click:
        //   rel:dx,dy,w,h|save=VarName|...
        // Also accepts shorthand:
        //   x,y,w,h|VarName
        var parts = (spec ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
            throw new Exception("OcrRead requires a spec like: x,y,w,h|save=VarName");

        var rectToken = parts[0].Trim();
        bool isRel = false;

        if (rectToken.StartsWith("rel:", StringComparison.OrdinalIgnoreCase) ||
            rectToken.StartsWith("last:", StringComparison.OrdinalIgnoreCase))
        {
            isRel = true;
            rectToken = rectToken.Substring(4);
        }

        var rectParts = rectToken.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (rectParts.Length != 4
            || !int.TryParse(rectParts[0], out var a)
            || !int.TryParse(rectParts[1], out var b)
            || !int.TryParse(rectParts[2], out var w)
            || !int.TryParse(rectParts[3], out var h))
            throw new Exception("OcrRead rect format must be: x,y,w,h or rel:dx,dy,w,h");

        int x = a, y = b;
        if (isRel)
        {
            if (lastClick == null)
                throw new Exception("OcrRead rel: requires a previous Click step (last click point is not set).");

            x = lastClick.Value.X + a;
            y = lastClick.Value.Y + b;
        }

        string save = "";
        string? regex = null;

        int retries = 1;
        int retryDelayMs = 120;
        bool preGray = false;
        int? thresh = null;
        bool invert = false;

        for (int i = 1; i < parts.Length; i++)
        {
            var p = parts[i];

            if (p.StartsWith("save=", StringComparison.OrdinalIgnoreCase))
            {
                save = p.Substring("save=".Length).Trim();
                continue;
            }
            if (p.StartsWith("regex=", StringComparison.OrdinalIgnoreCase))
            {
                regex = p.Substring("regex=".Length);
                continue;
            }
            if (p.StartsWith("retries=", StringComparison.OrdinalIgnoreCase) && int.TryParse(p.Substring(8), out var rr))
            {
                retries = Math.Clamp(rr, 1, 20);
                continue;
            }
            if ((p.StartsWith("retrydelay=", StringComparison.OrdinalIgnoreCase) || p.StartsWith("retryDelay=", StringComparison.OrdinalIgnoreCase)))
            {
                var val = p.Substring(p.IndexOf('=') + 1);
                if (int.TryParse(val, out var rd))
                    retryDelayMs = Math.Clamp(rd, 0, 5000);
                continue;
            }
            if (p.StartsWith("pre=", StringComparison.OrdinalIgnoreCase))
            {
                var mode = p.Substring("pre=".Length).Trim();
                if (mode.Equals("gray", StringComparison.OrdinalIgnoreCase) || mode.Equals("grayscale", StringComparison.OrdinalIgnoreCase))
                    preGray = true;
                continue;
            }
            if (p.StartsWith("thresh=", StringComparison.OrdinalIgnoreCase))
            {
                var val = p.Substring("thresh=".Length).Trim();
                if (int.TryParse(val, out var tv))
                    thresh = Math.Clamp(tv, 0, 255);
                continue;
            }
            if (p.StartsWith("invert=", StringComparison.OrdinalIgnoreCase))
            {
                var val = p.Substring("invert=".Length).Trim();
                invert = val == "1" || val.Equals("true", StringComparison.OrdinalIgnoreCase) || val.Equals("yes", StringComparison.OrdinalIgnoreCase);
                continue;
            }

            // Shorthand: second token is var name
            if (string.IsNullOrWhiteSpace(save))
                save = p.Trim();
        }

        if (string.IsNullOrWhiteSpace(save))
            throw new Exception("OcrRead requires save=VarName (or x,y,w,h|VarName shorthand)");

        var opt = new OcrWin.OcrOptions(preGray, thresh, invert);
        return new OcrSpec(new Rectangle(x, y, w, h), save, regex, retries, retryDelayMs, opt);
    }

    private readonly record struct SmartFindMatch(string MatchedText, Point ClickPoint, double Score);

    private static SmartFindMatch? TrySmartFindOnActiveWindow(
        string prompt,
        CancellationToken ct,
        double minScore = 0.60,
        bool allowFuzzy = true)
    {
        // Capture active window only (privacy + better search space)
        var cap = WindowCapture.CaptureForegroundWindowBitmap();
        using var bmp = cap.Bitmap;
        var winRect = cap.WindowRect;

        // OCR word boxes
        var boxes = OcrWin.RecognizeBitmapBoxesAsync(bmp, ct).GetAwaiter().GetResult();
        if (boxes.Count == 0)
            return null;

        var pNorm = NormalizeForMatch(prompt);
        if (pNorm.StartsWith("click ", StringComparison.Ordinal))
            pNorm = pNorm.Substring(6).Trim();

        SmartFindMatch? best = null;
        double bestScore = 0.0;

        foreach (var b in boxes)
        {
            ct.ThrowIfCancellationRequested();
            var t = b.Text;
            var tNorm = NormalizeForMatch(t);
            if (string.IsNullOrWhiteSpace(tNorm)) continue;

            double score = 0.0;

            if (tNorm == pNorm)
                score = 1.0;
            else if (WholeWordContains(tNorm, pNorm))
                score = 0.85;
            else if (tNorm.Contains(pNorm, StringComparison.Ordinal))
                score = 0.70;
            else if (allowFuzzy)
                score = 0.40 + 0.60 * SimilarityRatio(tNorm, pNorm); // keep fuzzy below exact

            if (score > bestScore)
            {
                // Convert box center to screen coords
                var cx = winRect.Left + b.Bounds.Left + (b.Bounds.Width / 2);
                var cy = winRect.Top + b.Bounds.Top + (b.Bounds.Height / 2);
                bestScore = score;
                best = new SmartFindMatch(t, new Point(cx, cy), score);
            }
        }

        if (best == null || best.Value.Score < minScore)
            return null;

        return best;
    }

    private static Dictionary<string, string> ParseKeyVals(string s)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in (s ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var eq = part.IndexOf('=');
            if (eq <= 0) continue;
            var k = part.Substring(0, eq).Trim();
            var v = part.Substring(eq + 1).Trim();
            if (!string.IsNullOrWhiteSpace(k))
                dict[k] = v;
        }
        return dict;
    }

    private static double ParseDouble(string? s, double fallback)
        => double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : fallback;

    private static bool ParseBool(string? s, bool fallback)
    {
        if (string.IsNullOrWhiteSpace(s)) return fallback;
        s = s.Trim();
        return s == "1" || s.Equals("true", StringComparison.OrdinalIgnoreCase) || s.Equals("yes", StringComparison.OrdinalIgnoreCase) || s.Equals("on", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeForMatch(string s)
    {
        s ??= "";
        s = s.Trim().ToLowerInvariant();
        var chars = new char[s.Length];
        int j = 0;
        bool lastSpace = false;
        for (int i = 0; i < s.Length; i++)
        {
            var c = s[i];
            if (char.IsLetterOrDigit(c))
            {
                chars[j++] = c;
                lastSpace = false;
            }
            else if (!lastSpace)
            {
                chars[j++] = ' ';
                lastSpace = true;
            }
        }
        return new string(chars, 0, j).Trim();
    }

    private static bool WholeWordContains(string text, string needle)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(needle)) return false;
        // Both are already normalized with spaces as separators
        var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            if (p == needle) return true;
        }
        return false;
    }

    // Levenshtein-based similarity ratio in [0..1]
    private static double SimilarityRatio(string a, string b)
    {
        if (a == b) return 1.0;
        if (a.Length == 0 || b.Length == 0) return 0.0;
        int dist = LevenshteinDistance(a, b);
        int max = Math.Max(a.Length, b.Length);
        return 1.0 - (double)dist / max;
    }

    private static int LevenshteinDistance(string a, string b)
    {
        int n = a.Length;
        int m = b.Length;
        var dp = new int[n + 1, m + 1];
        for (int i = 0; i <= n; i++) dp[i, 0] = i;
        for (int j = 0; j <= m; j++) dp[0, j] = j;
        for (int i = 1; i <= n; i++)
        {
            for (int j = 1; j <= m; j++)
            {
                int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
                int del = dp[i - 1, j] + 1;
                int ins = dp[i, j - 1] + 1;
                int sub = dp[i - 1, j - 1] + cost;
                dp[i, j] = Math.Min(Math.Min(del, ins), sub);
            }
        }
        return dp[n, m];
    }


    
    private static bool TryParseClickPoint(string? s, out int x, out int y)
    {
        x = y = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;
        var parts = s.Split(',', ';');
        if (parts.Length < 2) return false;
        return int.TryParse(parts[0].Trim(), out x) && int.TryParse(parts[1].Trim(), out y);
    }

private static string TruncateForLog(string s, int max = 80)
    {
        s ??= "";
        s = s.Replace("\r", " ").Replace("\n", " ").Trim();
        if (s.Length <= max) return s;
        return s.Substring(0, max) + "...";
    }

    private static void PasteText(string text)
    {
        // Clipboard requires STA. Runner may be called from a worker thread; handle both.
        try
        {
            SetClipboardTextSta(text ?? "");
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);
        }
        catch
        {
            // Fallback: direct typing
            Keyboard.Type(text ?? "");
        }
    }

    private static void SetClipboardTextSta(string text)
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
        {
            Clipboard.SetText(text);
            return;
        }

        Exception? error = null;
        var t = new Thread(() =>
        {
            try { Clipboard.SetText(text); }
            catch (Exception ex) { error = ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (error != null) throw error;
    }

    private static IDisposable? StartStopHotkeyPoller(RunnerOptions opts, Action onStop)
    {
        // Polling is reliable and avoids installing global keyboard hooks.
        // Uses the same GetAsyncKeyState approach as Recorder.
        var keyChar = char.ToUpperInvariant(opts.StopHotkeyKey);
        int vkey = keyChar;

        var timer = new System.Threading.Timer(_ =>
        {
            try
            {
                if (IsHotkeyPressed(opts, vkey))
                    onStop();
            }
            catch
            {
                // ignore polling errors
            }
        }, null, dueTime: 0, period: 50);

        return timer;
    }

    private static bool IsHotkeyPressed(RunnerOptions opts, int vkeyCode)
    {
        bool ctrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        bool shift = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
        bool alt = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
        bool keyDown = (GetAsyncKeyState(vkeyCode) & 0x8000) != 0;

        if (opts.StopHotkeyCtrl && !ctrl) return false;
        if (!opts.StopHotkeyCtrl && ctrl) return false;
        if (opts.StopHotkeyShift && !shift) return false;
        if (!opts.StopHotkeyShift && shift) return false;
        if (opts.StopHotkeyAlt && !alt) return false;
        if (!opts.StopHotkeyAlt && alt) return false;

        return keyDown;
    }

    private static void SleepWithStop(int ms, CancellationToken token)
    {
        if (ms <= 0) return;
        var end = Environment.TickCount64 + ms;
        while (Environment.TickCount64 < end)
        {
            token.ThrowIfCancellationRequested();
            Thread.Sleep(25);
        }
    }

    private const int VK_SHIFT = 0x10;
    private const int VK_CONTROL = 0x11;
    private const int VK_MENU = 0x12; // ALT

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    private static string? ReadText(AutomationElement? el)
    {
        if (el == null) return null;

        var tp = el.Patterns.Text.PatternOrDefault;
        if (tp != null)
        {
            var s = tp.DocumentRange.GetText(-1);
            return s?.Trim();
        }

        var vp = el.Patterns.Value.PatternOrDefault;
        if (vp != null)
        {
            var s = vp.Value.ValueOrDefault;
            return s?.Trim();
        }

        var name = el.Properties.Name.ValueOrDefault;
        return name?.Trim();
    }

    private static AutomationElement? Find(AutomationElement root, Target t)
    {
        var cf = root.Automation.ConditionFactory;
        ConditionBase? combined = null;

        void Add(ConditionBase c) => combined = combined == null ? c : combined.And(c);

        if (!string.IsNullOrWhiteSpace(t.AutomationId))
            Add(cf.ByAutomationId(t.AutomationId));
        if (!string.IsNullOrWhiteSpace(t.Name))
            Add(cf.ByName(t.Name));
        if (!string.IsNullOrWhiteSpace(t.ClassName))
            Add(cf.ByClassName(t.ClassName));
        if (!string.IsNullOrWhiteSpace(t.ControlType) &&
            Enum.TryParse<ControlType>(t.ControlType, out var ct))
            Add(cf.ByControlType(ct));

        if (combined == null) return null;
        return Wait(() => root.FindFirstDescendant(combined), TimeSpan.FromSeconds(6));
    }

    
    private static Point Click(AutomationElement el)
    {
        // We always compute a "best guess" center point for downstream features (like OCR relative to last click),
        // even if the click is performed via patterns.
        var r = el.BoundingRectangle;
        if (r.IsEmpty)
            throw new Exception("Cannot click element: empty BoundingRectangle");

        int cx = (int)(r.Left + r.Width / 2.0);
        int cy = (int)(r.Top + r.Height / 2.0);
        var center = new System.Drawing.Point(cx, cy);

        // Try InvokePattern first (fast/clean). If it fails (E_FAIL), fall back to a real mouse click.
        try
        {
            var invoke = el.Patterns.Invoke.PatternOrDefault;
            if (invoke != null)
            {
                invoke.Invoke();
                return center;
            }
        }
        catch
        {
            // ignore and fall back
        }

        try
        {
            // FlaUI's built-in Click can still use patterns; keep it but guard.
            el.Click();
            return center;
        }
        catch
        {
            // ignore and fall back
        }

        try
        {
            // Try focusing the element first
            el.Focus();
        }
        catch { }

        Mouse.Click(center);
        return center;
    }


    private static void SetText(AutomationElement el, string value)
    {
        var vp = el.Patterns.Value.PatternOrDefault;
        if (vp != null && !vp.IsReadOnly) { vp.SetValue(value); return; }

        el.Focus();
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
        Keyboard.Type(value);
    }

    private static void ClickImageMatch(ImageSearch.Match match, Target t)
    {
        var x = match.Location.X + (t.ClickOffsetX ?? (match.Size.Width / 2));
        var y = match.Location.Y + (t.ClickOffsetY ?? (match.Size.Height / 2));
        Mouse.Click(new Point(x, y));
    }

    private static string CreateRunDiagnosticsDir(string assetsDir)
    {
        try
        {
            // Prefer placing diagnostics next to the CSV (parent of *_assets) if possible.
            var root = assetsDir;
            try
            {
                var di = new DirectoryInfo(assetsDir);
                if (di.Name.EndsWith("_assets", StringComparison.OrdinalIgnoreCase) && di.Parent != null)
                    root = di.Parent.FullName;
            }
            catch { }

            var dir = Path.Combine(root, "_run_debug", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            Directory.CreateDirectory(dir);
            return dir;
        }
        catch
        {
            // As a last resort, use current directory.
            var dir = Path.Combine(Environment.CurrentDirectory, "_run_debug", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    private static void WriteStepDiagnostics(string runDir, int stepIndex, Step step, string? imgRel, Exception? findEx)
    {
        try
        {
            var prefix = $"step_{stepIndex + 1:000}_";

            // Screenshot before failing.
            using (var screen = ImageSearch.CaptureVirtualScreen())
            {
                var p = Path.Combine(runDir, prefix + "before.png");
                screen.Save(p, System.Drawing.Imaging.ImageFormat.Png);
            }

            // Copy template image if available.
            if (!string.IsNullOrWhiteSpace(imgRel))
            {
                // imgRel is relative to assets dir; if we cannot resolve, still write the value.
                File.WriteAllText(Path.Combine(runDir, prefix + "template_ref.txt"), imgRel);
            }

            // Basic context text.
            var txt = new List<string>
            {
                $"Action: {step.Action}",
                $"Note: {step.Note}",
                $"Value: {step.Value}",
                $"Target: {Describe(step.Target)}",
                $"Exception: {findEx?.GetType().Name}: {findEx?.Message}",
                $"Time: {DateTime.Now:O}",
                $"Foreground: {TryGetForegroundInfo()}"
            };
            File.WriteAllLines(Path.Combine(runDir, prefix + "info.txt"), txt);
        }
        catch
        {
            // never fail the run because diagnostics couldn't be written
        }
    }

    private static string TryGetForegroundInfo()
    {
        try
        {
            var h = Win32.GetForegroundWindow();
            if (h == IntPtr.Zero) return "(none)";
            if (!Win32.GetWindowRect(h, out var rect)) return "(no-rect)";

            var title = new System.Text.StringBuilder(512);
            _ = Win32.GetWindowText(h, title, title.Capacity);

            var w = rect.Right - rect.Left;
            var hgt = rect.Bottom - rect.Top;
            return $"'{title}' [{rect.Left},{rect.Top},{w}x{hgt}]";
        }
        catch { return "(unavailable)"; }
    }

    private static class Win32
    {
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        internal struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        internal static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    }

    private static Point? TryParsePoint(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var m = Regex.Match(value.Trim(), "^\\s*(?<x>-?\\d+)\\s*[, ]\\s*(?<y>-?\\d+)\\s*$");
        if (!m.Success) return null;
        if (!int.TryParse(m.Groups["x"].Value, out var x)) return null;
        if (!int.TryParse(m.Groups["y"].Value, out var y)) return null;
        return new Point(x, y);
    }

    private static T? Wait<T>(Func<T?> fn, TimeSpan timeout) where T : class
    {
        var sw = Stopwatch.StartNew();
        while (sw.Elapsed < timeout)
        {
            var v = fn();
            if (v != null) return v;
            Thread.Sleep(100);
        }
        return null;
    }

    private static string Describe(Target t)
        => $"AutomationId={t.AutomationId ?? "null"}, Name={t.Name ?? "null"}, ClassName={t.ClassName ?? "null"}, ControlType={t.ControlType ?? "null"}, Image={t.Image ?? "null"}";
    private static string GetText(AutomationElement? el)
    {
        if (el == null) return "";
        try
        {
            if (el.Patterns.Value.IsSupported)
                return el.Patterns.Value.Pattern.Value.Value ?? "";
        }
        catch { }

        try
        {
            // Some elements expose text via Name
            return el.Name ?? "";
        }
        catch { return ""; }
    }

}
