using System.Runtime.InteropServices;
using Avalonia.Controls;
using Avalonia.Input;
using RustPlusDesk.Core.Hotkeys;

namespace RustPlusDesk.Avalonia.Features.Hotkeys;

/// <summary>Where global hotkeys cannot be claimed: nothing registers and the reason is shown instead.</summary>
public sealed class UnavailableGlobalHotkeys(string reason) : IGlobalHotkeys
{
    public bool IsSupported => false;
    public string UnavailableReason => reason;
    public bool Register(string gesture) => false;
    public void UnregisterAll() { }
    public IReadOnlyDictionary<string, bool> RegistrationStatus { get; } = new Dictionary<string, bool>();
    public event Action<string>? Pressed { add { } remove { } }

    /// <summary>Truthful per-OS availability for this preview.</summary>
    public static UnavailableGlobalHotkeys ForCurrentPlatform() => new(OperatingSystem.IsMacOS()
        ? "Global hotkeys are not available on macOS in this preview yet (it needs the Accessibility permission flow). Assignments are saved and shared with the Windows app."
        : "Global hotkeys are not available on Linux in this preview yet: X11 and Wayland have no common global-shortcut API it supports. Assignments are saved and shared with the Windows app.");
}

/// <summary>
/// Reference GlobalHotkeyManager: RegisterHotKey on the main window with MOD_NOREPEAT, WM_HOTKEY routed back by ID,
/// and per-gesture registration status so the editor can flag keys another application already holds.
/// </summary>
public sealed class Win32GlobalHotkeys : IGlobalHotkeys, IDisposable
{
    private const uint ModAlt = 0x0001, ModControl = 0x0002, ModShift = 0x0004, ModWin = 0x0008, ModNoRepeat = 0x4000;
    private const uint WmHotkey = 0x0312;
    private readonly IntPtr hwnd;
    private readonly TopLevel window;
    private readonly Win32Properties.CustomWndProcHookCallback hook;
    private int nextId = 1;
    private readonly Dictionary<int, string> idToGesture = [];
    private readonly Dictionary<string, int> gestureToId = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, bool> status = new(StringComparer.OrdinalIgnoreCase);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public Win32GlobalHotkeys(TopLevel window)
    {
        this.window = window;
        hwnd = window.TryGetPlatformHandle()?.Handle ?? IntPtr.Zero;
        hook = OnMessage;
        Win32Properties.AddWndProcHookCallback(window, hook);
    }

    public bool IsSupported => hwnd != IntPtr.Zero;
    public string UnavailableReason => IsSupported ? "" : "The window handle is not available yet.";
    public IReadOnlyDictionary<string, bool> RegistrationStatus => status;
    public event Action<string>? Pressed;

    public bool Register(string gesture)
    {
        if (!TryParse(gesture, out uint modifiers, out uint vk)) { status[gesture] = false; return false; }
        if (gestureToId.ContainsKey(gesture)) return status.GetValueOrDefault(gesture);
        int id = nextId++;
        bool ok = RegisterHotKey(hwnd, id, modifiers | ModNoRepeat, vk);
        status[gesture] = ok;
        if (ok) { gestureToId[gesture] = id; idToGesture[id] = gesture; }
        return ok;
    }

    public void UnregisterAll()
    {
        foreach (int id in idToGesture.Keys) UnregisterHotKey(hwnd, id);
        idToGesture.Clear();
        gestureToId.Clear();
        status.Clear();
    }

    public void Dispose()
    {
        UnregisterAll();
        Win32Properties.RemoveWndProcHookCallback(window, hook);
    }

    private IntPtr OnMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey && idToGesture.TryGetValue(wParam.ToInt32(), out var gesture))
        {
            handled = true;
            Pressed?.Invoke(gesture);
        }
        return IntPtr.Zero;
    }

    internal static bool TryParse(string gesture, out uint modifiers, out uint vk)
    {
        modifiers = 0;
        vk = 0;
        if (!HotkeyGesture.TryParse(gesture, out bool ctrl, out bool alt, out bool shift, out bool win, out string keyName)) return false;
        if (!Enum.TryParse(keyName, true, out Key key) || VirtualKey(key) is not { } code) return false;
        if (ctrl) modifiers |= ModControl;
        if (alt) modifiers |= ModAlt;
        if (shift) modifiers |= ModShift;
        if (win) modifiers |= ModWin;
        vk = code;
        return true;
    }

    /// <summary>Windows virtual-key codes for the keys a hotkey can use.</summary>
    internal static uint? VirtualKey(Key key) => key switch
    {
        >= Key.A and <= Key.Z => (uint)(0x41 + (key - Key.A)),
        >= Key.D0 and <= Key.D9 => (uint)(0x30 + (key - Key.D0)),
        >= Key.F1 and <= Key.F24 => (uint)(0x70 + (key - Key.F1)),
        >= Key.NumPad0 and <= Key.NumPad9 => (uint)(0x60 + (key - Key.NumPad0)),
        Key.Multiply => 0x6A, Key.Add => 0x6B, Key.Subtract => 0x6D, Key.Decimal => 0x6E, Key.Divide => 0x6F,
        Key.Space => 0x20, Key.Enter => 0x0D, Key.Tab => 0x09, Key.Back => 0x08, Key.Escape => 0x1B,
        Key.Insert => 0x2D, Key.Delete => 0x2E, Key.Home => 0x24, Key.End => 0x23, Key.PageUp => 0x21, Key.PageDown => 0x22,
        Key.Left => 0x25, Key.Up => 0x26, Key.Right => 0x27, Key.Down => 0x28,
        Key.Pause => 0x13, Key.Scroll => 0x91, Key.PrintScreen => 0x2C, Key.CapsLock => 0x14, Key.NumLock => 0x90,
        Key.OemSemicolon => 0xBA, Key.OemPlus => 0xBB, Key.OemComma => 0xBC, Key.OemMinus => 0xBD, Key.OemPeriod => 0xBE,
        Key.OemQuestion => 0xBF, Key.OemTilde => 0xC0, Key.OemOpenBrackets => 0xDB, Key.OemPipe => 0xDC,
        Key.OemCloseBrackets => 0xDD, Key.OemQuotes => 0xDE, Key.OemBackslash => 0xE2,
        _ => null
    };
}
