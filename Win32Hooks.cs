using System.Runtime.InteropServices;
using System.Text;

namespace PortableWinFormsRecorder;

/// <summary>
/// Dependency-free global low-level mouse + keyboard hooks (WH_MOUSE_LL / WH_KEYBOARD_LL).
/// </summary>
public static class Win32Hooks
{
    public enum MouseButton { Left, Right, Middle, X1, X2 }
    public sealed record MouseEvent(int X, int Y, MouseButton Button);
    public sealed record KeyEvent(int VkCode, bool Ctrl, bool Shift, bool Alt);

    public static event Action<MouseEvent>? MouseDown;
    public static event Action<KeyEvent>? KeyDown;
    public static event Action<char>? KeyPress;

    private const int WH_MOUSE_LL = 14;
    private const int WH_KEYBOARD_LL = 13;

    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_XBUTTONDOWN = 0x020B;

    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;

    private static IntPtr _mouseHook = IntPtr.Zero;
    private static IntPtr _kbdHook = IntPtr.Zero;

    private static LowLevelMouseProc? _mouseProc;
    private static LowLevelKeyboardProc? _kbdProc;

    public static bool IsRunning => _mouseHook != IntPtr.Zero && _kbdHook != IntPtr.Zero;

    public static void Start()
    {
        if (IsRunning) return;

        _mouseProc = MouseHookCallback;
        _kbdProc = KeyboardHookCallback;

        using var curProcess = System.Diagnostics.Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule!;
        IntPtr hMod = IntPtr.Zero;

        _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, hMod, 0);
        _kbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, _kbdProc, hMod, 0);

        if (_mouseHook == IntPtr.Zero || _kbdHook == IntPtr.Zero)
        {
            Stop();
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Failed to set Windows hook");
        }
    }

    public static void Stop()
    {
        if (_mouseHook != IntPtr.Zero) { UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
        if (_kbdHook != IntPtr.Zero) { UnhookWindowsHookEx(_kbdHook); _kbdHook = IntPtr.Zero; }
        _mouseProc = null;
        _kbdProc = null;
    }

    private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            int msg = wParam.ToInt32();
            var info = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

            MouseButton? btn = msg switch
            {
                WM_LBUTTONDOWN => MouseButton.Left,
                WM_RBUTTONDOWN => MouseButton.Right,
                WM_MBUTTONDOWN => MouseButton.Middle,
                WM_XBUTTONDOWN => ((info.mouseData >> 16) & 0xffff) == 1 ? MouseButton.X1 : MouseButton.X2,
                _ => null
            };

            if (btn != null)
                MouseDown?.Invoke(new MouseEvent(info.pt.x, info.pt.y, btn.Value));
        }

        return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    private static IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            int msg = wParam.ToInt32();
            var info = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);

            if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN)
            {
                bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;

                KeyDown?.Invoke(new KeyEvent(info.vkCode, ctrl, shift, alt));

                // Best-effort KeyPress (printable chars)
                char? ch = TryToChar(info.vkCode, info.scanCode);
                if (ch != null && !char.IsControl(ch.Value))
                    KeyPress?.Invoke(ch.Value);
            }
        }

        return CallNextHookEx(_kbdHook, nCode, wParam, lParam);
    }

    private static char? TryToChar(int vkCode, int scanCode)
    {
        byte[] state = new byte[256];
        if (!GetKeyboardState(state)) return null;

        var sb = new StringBuilder(8);
        int rc = ToUnicode((uint)vkCode, (uint)scanCode, state, sb, sb.Capacity, 0);
        if (rc == 1) return sb[0];
        return null;
    }

    private const int VK_SHIFT = 0x10;
    private const int VK_CONTROL = 0x11;
    private const int VK_MENU = 0x12;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int x; public int y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public int mouseData;
        public int flags;
        public int time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public int vkCode;
        public int scanCode;
        public int flags;
        public int time;
        public IntPtr dwExtraInfo;
    }

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int nVirtKey);

    [DllImport("user32.dll")]
    private static extern bool GetKeyboardState(byte[] lpKeyState);

    [DllImport("user32.dll")]
    private static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState,
        [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags);
}