using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace PortableWinFormsRecorder;

public sealed class RecorderOptions
{
    public bool VarMode { get; set; }
    public bool CaptureImages { get; set; }
    public string AssetsDir { get; set; } = ".";

    // Stop hotkey configuration (global, via polling)
    public bool StopHotkeyCtrl { get; set; } = true;
    public bool StopHotkeyShift { get; set; } = true;
    public bool StopHotkeyAlt { get; set; } = false;
    public char StopHotkeyKey { get; set; } = 'S';

    // Pause/Resume hotkey (default Ctrl+Shift+P)
    public bool PauseHotkeyCtrl { get; set; } = true;
    public bool PauseHotkeyShift { get; set; } = true;
    public bool PauseHotkeyAlt { get; set; } = false;
    public char PauseHotkeyKey { get; set; } = 'P';
}

public sealed class Recorder : IDisposable
{
    public event Action<bool>? PausedChanged;
    public bool IsPaused { get; private set; }
    private bool _lastPauseCombo;

    private long _ignoreInputUntilTick;

    private readonly string _processName;
    private readonly Script _script;
    private readonly RecorderOptions _opts;

    private volatile bool _stopRequested;
    private volatile bool _isStopping;

    // Fallback polling (works even if hooks are blocked by policy)
    private System.Threading.Timer? _pollTimer;
    private bool _lastLeftDown;
    private bool _lastStopCombo;
    private long _lastAnyClickTick;
    private long _lastHookClickTick;
    private int _lastClickX;
    private int _lastClickY;

    private UIA3Automation? _automation;
    private FlaUI.Core.Application? _app;
    private FlaUI.Core.AutomationElements.Window? _window;

    // When typing is captured we try to bind it to a focused editable element.
    // If we cannot reliably identify an editable element (common in modern Notepad / custom editors),
    // we still record a TypeText step (no target) so the runner can replay keystrokes into the focused app.
    private Target? _activeEditTarget;
    private bool _typedWithoutTarget;
    private readonly StringBuilder _typed = new();
    private DateTime _lastKey = DateTime.MinValue;

    private int _stepCounter = 0;

    public event Action<string>? Log;
    public event Action<Step, string>? StepRecorded;

    public Recorder(string processName, Script script, RecorderOptions opts)
    {
        _processName = processName;
        _script = script;
        _opts = opts;
    }

    public bool IsRecording { get; private set; }

    public void Start()
    {
        _ignoreInputUntilTick = Environment.TickCount64 + 250; // Ignore initial click-through // IgnoreInputUntil

        
        
        _lastLeftDown = false;
_lastAnyClickTick = 0;
        _lastHookClickTick = 0;
        _lastClickX = 0;
        _lastClickY = 0;
_stopRequested = false;
        _isStopping = false;
        _automation = new UIA3Automation();

        if (!string.IsNullOrWhiteSpace(_processName))
        {
            var proc = Process.GetProcessesByName(_processName).FirstOrDefault()
                       ?? throw new Exception($"Process not found: {_processName}");

            _app = FlaUI.Core.Application.Attach(proc);
            _window = _app.GetMainWindow(_automation) ?? throw new Exception("Main window not found");
            _window.Focus();
            Log?.Invoke($"[REC] Attached to process: {_processName}");
        }
        else
        {
            _window = null; // Desktop mode
            Log?.Invoke("[REC] Desktop mode (no process). Recording globally.");
        }

        Win32Hooks.Start();
        Win32Hooks.MouseDown += OnMouseDown;
        Win32Hooks.KeyPress += OnKeyPress;
        Win32Hooks.KeyDown += OnKeyDown;

        Log?.Invoke("[REC] hooks attached (Win32 low-level mouse+keyboard)");
        StartPolling();

        IsRecording = true;
        Log?.Invoke("[REC] started");
    }

    public void TogglePauseFromUi()
    {
        TogglePause();
    }

    private void TogglePause()
    {
        IsPaused = !IsPaused;
        try { PausedChanged?.Invoke(IsPaused); } catch { }
        try { Log?.Invoke(IsPaused ? "[REC] paused" : "[REC] resumed"); } catch { }
    }

    public void RequestStop()
    {
        _stopRequested = true;
        _isStopping = true;
        IsRecording = false;

        try
        {
            _pollTimer?.Dispose();
            _pollTimer = null;

            Win32Hooks.MouseDown -= OnMouseDown;
            Win32Hooks.KeyPress -= OnKeyPress;
            Win32Hooks.KeyDown -= OnKeyDown;
            Win32Hooks.Stop();
        }
        catch { }

        Log?.Invoke("[REC] stop requested");
    }

    public void WaitUntilStopped()
    {
        while (!_stopRequested)
        {
            FlushIfIdle();
            Thread.Sleep(50);
        }

        Flush(force: true);
        IsRecording = false;
        Log?.Invoke("[REC] stopped");
    }

    private void OnMouseDown(Win32Hooks.MouseEvent e)
    {
        if (IsPaused) return;

        if (Environment.TickCount64 < _ignoreInputUntilTick) return;

        if (!IsRecording || _isStopping) return;

        _lastHookClickTick = Environment.TickCount64;

        // Debounce duplicate click events (hooks + polling can overlap)
        var now = _lastHookClickTick;
        if (now - _lastAnyClickTick < CLICK_DEBOUNCE_MS)
        {
            if (Math.Abs(e.X - _lastClickX) <= CLICK_DEBOUNCE_PX && Math.Abs(e.Y - _lastClickY) <= CLICK_DEBOUNCE_PX)
                return;
        }
        _lastAnyClickTick = now;
        _lastClickX = e.X;
        _lastClickY = e.Y;

        if (e.Button != Win32Hooks.MouseButton.Left) return;
        if (_lastLeftDown) return; // ignore repeats while button is held
        _lastLeftDown = true;

        Log?.Invoke($"[REC] mouse event at {e.X},{e.Y}");

        Flush(force: true);
        
        var pt = new Point(e.X, e.Y);
        var el = (_window != null) ? _window.Automation.FromPoint(pt) : _automation!.FromPoint(pt);
        if (el == null) return;
        // In process mode, ignore clicks outside that process. In desktop mode, accept all.
        if (_window != null)
        {
            var pid = el.Properties.ProcessId.ValueOrDefault;
            var targetPid = _window.Properties.ProcessId.ValueOrDefault;
            if (pid != targetPid) return;
        }

        var target = Selectors.FromElement(el, _window?.Properties.Name.ValueOrDefault);

        if (_opts.CaptureImages || Selectors.IsTooWeak(target))
        {
            var rel = TryCaptureElementImage(el);
            if (!string.IsNullOrWhiteSpace(rel))
                target.Image = rel;
        }

        _script.Steps.Add(new Step { Action = "Click", Target = target });
        // Flow image for UI timeline
        var flowImg = TryCaptureFlowImage(el);
        if (!string.IsNullOrWhiteSpace(flowImg))
        {
            var stepRef = _script.Steps.Last();
            // Store flow image filename for CSV/Flow editor
            if (string.IsNullOrWhiteSpace(stepRef.Target.Image))
                stepRef.Target.Image = Path.GetFileName(flowImg);
            StepRecorded?.Invoke(stepRef, flowImg);
        }

        Log?.Invoke($"[REC] Click: {Describe(target)}");

        // Many editors (e.g., modern Notepad / custom text areas) expose ControlType.Document instead of Edit.
        // Treat both as "editable" to enable text capture.
        _typedWithoutTarget = false;

        if (IsEditableTarget(target))
        {
            _activeEditTarget = target;
            _typed.Clear();
        }
        else
        {
            _activeEditTarget = null;
            _typed.Clear();
        }
    }

    private void OnKeyPress(char ch)
    {
        if (IsPaused) return;

        if (Environment.TickCount64 < _ignoreInputUntilTick) return;

        if (!IsRecording || _isStopping) return;

        if (_activeEditTarget == null && !_typedWithoutTarget)
        {
            // If we didn't click an "editable" element first, try to infer it from the currently focused element.
            // If we can't prove it's editable (common in modern Notepad/custom editors), we still capture typing
            // as a TypeText step (no target) so replay works.
            var inferred = TryGetFocusedTarget(out var editable);
            if (inferred != null)
            {
                if (editable)
                {
                    _activeEditTarget = inferred;
                    _typedWithoutTarget = false;
                }
                else
                {
                    _activeEditTarget = null;
                    _typedWithoutTarget = true;
                }
                _typed.Clear();
            }
        }

        if (_activeEditTarget == null && !_typedWithoutTarget) return;
        if (char.IsControl(ch)) return;

        _typed.Append(ch);
        _lastKey = DateTime.UtcNow;
    }

    private void OnKeyDown(Win32Hooks.KeyEvent e)
    {
        if (IsPaused) return;

        if (Environment.TickCount64 < _ignoreInputUntilTick) return;

        if (!IsRecording || _isStopping) return;

        // Stop: Ctrl+Shift+S
        if (e.Ctrl && e.Shift && e.VkCode == 0x53)
        {
            _stopRequested = true;
        _isStopping = true;
            Log?.Invoke("[REC] stop requested (hotkey)");
            return;
        }

        if (_activeEditTarget == null && !_typedWithoutTarget)
        {
            // Keep trying to bind typed keys to the focused target.
            var inferred = TryGetFocusedTarget(out var editable);
            if (inferred != null && editable)
                _activeEditTarget = inferred;
        }

        if ((_activeEditTarget != null || _typedWithoutTarget) && e.VkCode == 0x08 && _typed.Length > 0)
        {
            _typed.Length -= 1;
            _lastKey = DateTime.UtcNow;
        }

        if ((_activeEditTarget != null || _typedWithoutTarget) && (e.VkCode == 0x0D || e.VkCode == 0x09))
            Flush(force: true);
    }

    private static bool IsEditableTarget(Target t)
    {
        var ct = t.ControlType ?? string.Empty;
        return ct.Equals(ControlType.Edit.ToString(), StringComparison.OrdinalIgnoreCase)
               || ct.Equals(ControlType.Document.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    private Target? TryGetFocusedTarget(out bool editable)
    {
        editable = false;
        try
        {
            if (_automation == null) return null;

            // UIA3Automation exposes FocusedElement() as a method in this FlaUI version.
            // Use it to obtain the current focused element.
            var focused = _automation.FocusedElement();
            if (focused == null) return null;

            // In process mode, ignore other processes.
            if (_window != null)
            {
                var pid = focused.Properties.ProcessId.ValueOrDefault;
                var targetPid = _window.Properties.ProcessId.ValueOrDefault;
                if (pid != targetPid) return null;
            }

            editable = IsProbablyEditable(focused);
            var target = Selectors.FromElement(focused, _window?.Properties.Name.ValueOrDefault);
            // If it's editable by patterns we accept it even if ControlType isn't Edit/Document.
            if (!editable && IsEditableTarget(target)) editable = true;
            return target;
        }
        catch { return null; }
    }

    private static bool IsProbablyEditable(AutomationElement el)
    {
        try
        {
            // Prefer pattern support over ControlType (modern Notepad often uses custom elements).
            var p = el.Patterns;
            if (p.Value.IsSupported) return true;
            if (p.Text.IsSupported) return true;
        }
        catch { }
        return false;
    }

    private void FlushIfIdle()
    {
        if ((_activeEditTarget == null && !_typedWithoutTarget) || _typed.Length == 0) return;
        if ((DateTime.UtcNow - _lastKey).TotalMilliseconds > 800)
            Flush(force: true);
    }

    private void Flush(bool force)
    {
        if (_activeEditTarget == null && !_typedWithoutTarget) return;
        if (_typed.Length == 0) { _activeEditTarget = null; _typedWithoutTarget = false; return; }

        var captured = _typed.ToString();
        var value = captured;

        if (_opts.VarMode)
        {
            Log?.Invoke($"[REC] Captured text: \"{captured}\"");
            Console.Write("Replace with variable? (enter var name or blank): ");
            var name = Console.ReadLine()?.Trim();
            if (!string.IsNullOrWhiteSpace(name))
                value = "{{" + name + "}}";
        }

        if (_activeEditTarget != null)
        {
            _script.Steps.Add(new Step { Action = "SetText", Target = _activeEditTarget, Value = value });
            Log?.Invoke($"[REC] SetText: {Describe(_activeEditTarget)} = {value}");
        }
        else
        {
            _script.Steps.Add(new Step { Action = "TypeText", Value = value });
            Log?.Invoke($"[REC] TypeText: {value}");
        }

        _typed.Clear();
        _activeEditTarget = null;
        _typedWithoutTarget = false;
    }

    private string? TryCaptureElementImage(AutomationElement el)
    {
        try
        {
            var r = el.BoundingRectangle;
            if (r.IsEmpty) return null;

            int x = (int)Math.Round((double)r.Left);
            int y = (int)Math.Round((double)r.Top);
            int w = (int)Math.Round((double)r.Width);
            int h = (int)Math.Round((double)r.Height);
            if (w < 8 || h < 8) return null;

            int pad = 2;
            x -= pad; y -= pad; w += pad * 2; h += pad * 2;

            using var bmp = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bmp))
                g.CopyFromScreen(x, y, 0, 0, new Size(w, h));

            _stepCounter++;
            var file = $"step_{_stepCounter:0000}.png";
            var full = Path.Combine(_opts.AssetsDir, file);
            bmp.Save(full, ImageFormat.Png);

            return file;
        }
        catch { return null; }
    }

    private string? TryCaptureFlowImage(FlaUI.Core.AutomationElements.AutomationElement el)
    {
        try
        {
            var r = el.BoundingRectangle;
            if (r.IsEmpty) return null;

            var vs = SystemInformation.VirtualScreen;
            using var bmp = new Bitmap(vs.Width, vs.Height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(vs.Left, vs.Top, 0, 0, vs.Size);

                // draw highlight
                using var pen = new Pen(Color.Red, 4);
                var rect = new Rectangle(
                    (int)((double)r.Left - vs.Left),
                    (int)((double)r.Top - vs.Top),
                    (int)((double)r.Width),
                    (int)((double)r.Height)
                );
                g.DrawRectangle(pen, rect);
            }

            _stepCounter++;
            var file = $"flow_{_stepCounter:0000}.png";
            var full = Path.Combine(_opts.AssetsDir, file);
            bmp.Save(full, ImageFormat.Png);
            return full;
        }
        catch { return null; }
    }

    private static string Describe(Target t)
        => $"AutomationId={t.AutomationId ?? "null"}, Name={t.Name ?? "null"}, ClassName={t.ClassName ?? "null"}, ControlType={t.ControlType ?? "null"}, Image={t.Image ?? "null"}";

    private void StartPolling()
    {
        // 50 Hz polling for left button transitions
        _lastLeftDown = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
        _lastStopCombo = false;
        _pollTimer = new System.Threading.Timer(_ =>
        {
            try
            {
                var down = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
                if (down && !_lastLeftDown)
                {
                    var now = Environment.TickCount64;
                    if (_lastHookClickTick != 0 && (now - _lastHookClickTick) < HOOK_PREFERRED_WINDOW_MS)
                    {
                        // hook already captured the click
                    }
                    else
                    {
                        var p = Cursor.Position;
                        // Use the same handler; debounce will prevent repeats
                        OnMouseDown(new Win32Hooks.MouseEvent(p.X, p.Y, Win32Hooks.MouseButton.Left));
                        _lastLeftDown = true;
                    }
                }
                _lastLeftDown = down;

                // Ctrl+Shift+S stop hotkey (polling fallback)
                var ctrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                var shift = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
                var alt = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;

                var keyChar = _opts.StopHotkeyKey;
                var vkey = char.IsLetterOrDigit(keyChar) ? char.ToUpperInvariant(keyChar) : 'S';
                var vkeyCode = (int)vkey;
                var keyDown = (GetAsyncKeyState(vkeyCode) & 0x8000) != 0;

                var combo = (!_opts.StopHotkeyCtrl || ctrl)
                            && (!_opts.StopHotkeyShift || shift)
                            && (!_opts.StopHotkeyAlt || alt)
                            && keyDown;

                if (combo && !_lastStopCombo)
                {
                    _lastStopCombo = true;
                    Log?.Invoke($"[REC] Stop hotkey detected ({FormatHotkey()})");
                    RequestStop();
                    return;
                }
                if (!combo) _lastStopCombo = false;

                // Ctrl+Shift+P pause/resume hotkey
                var pKeyChar = _opts.PauseHotkeyKey;
                var pVkey = char.IsLetterOrDigit(pKeyChar) ? char.ToUpperInvariant(pKeyChar) : 'P';
                var pVkeyCode = (int)pVkey;
                var pKeyDown = (GetAsyncKeyState(pVkeyCode) & 0x8000) != 0;

                var pCombo = (!_opts.PauseHotkeyCtrl || ctrl)
                             && (!_opts.PauseHotkeyShift || shift)
                             && (!_opts.PauseHotkeyAlt || alt)
                             && pKeyDown;

                if (pCombo && !_lastPauseCombo)
                {
                    _lastPauseCombo = true;
                    Log?.Invoke($"[REC] Pause hotkey detected (Ctrl+Shift+{pVkey})");
                    TogglePause();
                }
                if (!pCombo) _lastPauseCombo = false;

            }
            catch { /* swallow */ }
        }, null, 0, 20);
        Log?.Invoke("[REC] polling fallback enabled (GetAsyncKeyState)");
    }

    private const int VK_LBUTTON = 0x01;
    private const int VK_SHIFT = 0x10;
    private const int VK_CONTROL = 0x11;
    private const int VK_MENU = 0x12; // ALT key
    private const int VK_S = 0x53;
    private const int CLICK_DEBOUNCE_MS = 180;
    private const int CLICK_DEBOUNCE_PX = 3;
    private const int HOOK_PREFERRED_WINDOW_MS = 250;

    
    private string FormatHotkey()
    {
        var parts = new List<string>();
        if (_opts.StopHotkeyCtrl) parts.Add("Ctrl");
        if (_opts.StopHotkeyShift) parts.Add("Shift");
        if (_opts.StopHotkeyAlt) parts.Add("Alt");
        parts.Add(char.ToUpperInvariant(_opts.StopHotkeyKey).ToString());
        return string.Join("+", parts);
    }

[System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    public void Dispose()
    {
                try
        {
            _pollTimer?.Dispose();
            _pollTimer = null;

            Win32Hooks.MouseDown -= OnMouseDown;
            Win32Hooks.KeyPress -= OnKeyPress;
            Win32Hooks.KeyDown -= OnKeyDown;
            Win32Hooks.Stop();
        }
        catch { }

        try { _app?.Dispose(); } catch { }
        try { _automation?.Dispose(); } catch { }
    }
}
