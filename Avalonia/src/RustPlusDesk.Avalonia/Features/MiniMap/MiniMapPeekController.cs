using System.Runtime.InteropServices;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using RustPlusDesk.Core.Hotkeys;

namespace RustPlusDesk.Avalonia.Features.MiniMap;

/// <summary>
/// Reference MainWindow.MiniMapPeek: hold a key, see the map; let go, it is gone.
///
/// A second monitor is the thing that makes a mini-map comfortable, and plenty of people have one screen. Leaving it
/// up over the game on that screen is an advantage the game does not give anyone - Rust's own map takes a keypress and
/// covers what you were looking at, and that cost is the point of it.
///
/// So this gives the same bargain: the map is only on screen while the key is held, over the game, the way the in-game
/// map is. Nobody gets to watch the map and the world at once.
///
/// The bargain has a time limit as well. After twenty seconds the map goes away whether the key is still down or not,
/// and stays away until the key is released and pressed again: taping the key down and playing on with the map in the
/// corner is not that bargain.
///
/// Windows only tells the app that a key went down, so the release is tracked with a low-level keyboard hook that is
/// active only while the map is up, plus the platform's own key state as a backup.
/// </summary>
public sealed class MiniMapPeekController : IDisposable
{
    private const int WhKeyboardLl = 13;
    private const int WmKeyDown = 0x0100, WmKeyUp = 0x0101, WmSysKeyDown = 0x0104, WmSysKeyUp = 0x0105;
    private const uint ModNoRepeat = 0x4000;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardHookStruct
    {
        public uint VirtualKey;
        public uint ScanCode;
        public uint Flags;
        public uint Time;
        public UIntPtr ExtraInfo;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint threadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? moduleName);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int virtualKey);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint virtualKey);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterHotKey(IntPtr window, int id);

    /// <summary>How long the map can be held up for.</summary>
    public static readonly TimeSpan Limit = TimeSpan.FromSeconds(20);

    private readonly TopLevel window;
    private readonly Win32Properties.CustomWndProcHookCallback messageHook;
    private LowLevelKeyboardProc? keyboardHook;
    private readonly DispatcherTimer limitTimer;

    private IntPtr handle;
    private IntPtr keyboardHookId;
    private uint peekVirtualKey;
    private int registeredId;
    private string gesture = "";
    private bool peeking;
    private bool spent;

    /// <summary>The map is shown while the key is held.</summary>
    public Action? Show { get; set; }

    /// <summary>The map is put away again.</summary>
    public Action? Hide { get; set; }

    /// <summary>Whether the map is up right now.</summary>
    public bool IsPeeking => peeking;

    /// <summary>
    /// Whether the peek works on this machine.
    ///
    /// The release of a key is only told to a Windows application, so the peek is a Windows feature in this port. The
    /// setting says so where it is offered rather than quietly doing nothing.
    /// </summary>
    public bool IsSupported => OperatingSystem.IsWindows();

    /// <summary>Why the peek cannot be used here, when it cannot.</summary>
    public string UnavailableReason => OperatingSystem.IsWindows()
        ? ""
        : "Holding a key to see the map over the game needs the Windows keyboard hook, which the other platforms do not have yet.";

    public MiniMapPeekController(TopLevel window)
    {
        this.window = window;
        handle = window.TryGetPlatformHandle()?.Handle ?? IntPtr.Zero;
        messageHook = OnMessage;
        Win32Properties.AddWndProcHookCallback(window, messageHook);

        limitTimer = new DispatcherTimer { Interval = Limit };
        limitTimer.Tick += (_, _) => EndPeek(spent: true);
    }

    /// <summary>
    /// The key that shows the map, or null for none.
    ///
    /// The gesture is the reference's own form ("Ctrl+Alt+K" or a plain key), and the key that ends the peek is the
    /// one in it.
    /// </summary>
    public string Gesture
    {
        get => gesture;
        set
        {
            string? wanted = string.IsNullOrWhiteSpace(value) ? null : value.Trim();
            if (string.Equals(gesture, wanted, StringComparison.OrdinalIgnoreCase)) return;

            Unregister();
            gesture = wanted ?? "";
            Register();
        }
    }

    /// <summary>Claims the key, so the map comes up when it is pressed.</summary>
    private void Register()
    {
        if (!IsSupported || string.IsNullOrEmpty(gesture) || handle == IntPtr.Zero) return;
        if (!Features.Hotkeys.Win32GlobalHotkeys.TryParse(gesture, out uint modifiers, out uint virtualKey)) return;

        if (!HotkeyGesture.TryParse(gesture, out bool ctrl, out bool alt, out bool shift, out bool win, out string keyName)) return;
        if (!Enum.TryParse(keyName, true, out Key key)) return;
        if (Features.Hotkeys.Win32GlobalHotkeys.VirtualKey(key) is not { } code) return;

        peekVirtualKey = code;
        registeredId = GetHashCode() & 0x7FFF;
        if (registeredId == 0) registeredId = 1;

        if (!RegisterHotKey(handle, registeredId, modifiers | ModNoRepeat, virtualKey)) registeredId = 0;
    }

    /// <summary>Gives the key back.</summary>
    private void Unregister()
    {
        if (registeredId != 0 && handle != IntPtr.Zero) UnregisterHotKey(handle, registeredId);
        registeredId = 0;
    }

    /// <summary>The window's messages: the key that starts the peek.</summary>
    private IntPtr OnMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        const uint WmHotkey = 0x0312;
        if (msg == WmHotkey && registeredId != 0 && wParam.ToInt32() == registeredId)
        {
            handled = true;
            BeginPeek();
        }
        return IntPtr.Zero;
    }

    /// <summary>The map comes up, and the release is watched for.</summary>
    public void BeginPeek()
    {
        if (peeking) return;

        // A hold that has already run out of time must not simply start again: the key is auto-repeating into this
        // while it is held.
        if (spent) return;
        if (peekVirtualKey == 0) return;

        peeking = true;
        InstallKeyboardHook();
        limitTimer.Start();
        Show?.Invoke();
    }

    /// <summary>The map goes away.</summary>
    public void EndPeek(bool spent = false)
    {
        limitTimer.Stop();
        RemoveKeyboardHook();

        if (!peeking) return;
        peeking = false;
        this.spent = spent;
        Hide?.Invoke();

        // The key that was held is released at some point; the next hold is a new one.
        if (!spent) spent = false;
    }

    /// <summary>
    /// The low-level hook, which is the only way to hear that a key was let go.
    ///
    /// It is installed only while the map is up, and only on Windows.
    /// </summary>
    private void InstallKeyboardHook()
    {
        if (!IsSupported || keyboardHookId != IntPtr.Zero) return;

        IntPtr module = GetModuleHandle(null);
        keyboardHookId = SetWindowsHookEx(WhKeyboardLl, keyboardHook ??= OnKeyboard, module, 0);
    }

    /// <summary>Takes the hook down again.</summary>
    private void RemoveKeyboardHook()
    {
        if (keyboardHookId == IntPtr.Zero) return;
        UnhookWindowsHookEx(keyboardHookId);
        keyboardHookId = IntPtr.Zero;
    }

    /// <summary>The hook's own words: the key up ends the peek, and the platform's key state is the backup.</summary>
    private IntPtr OnKeyboard(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            int message = wParam.ToInt32();
            bool down = message is WmKeyDown or WmSysKeyDown;
            bool up = message is WmKeyUp or WmSysKeyUp;

            if (up && peekVirtualKey != 0)
            {
                var pressed = Marshal.PtrToStructure<KeyboardHookStruct>(lParam);
                if (pressed.VirtualKey == peekVirtualKey) EndPeek();
            }
            else if (down && peekVirtualKey != 0 && (GetAsyncKeyState((int)peekVirtualKey) & 0x8000) == 0)
            {
                // The key was let go while the app was not looking - the window lost the focus, or the hook was
                // installed a moment too late. The map goes away rather than staying up on its own.
                EndPeek();
            }
        }

        return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
    }

    public void Dispose()
    {
        limitTimer.Stop();
        RemoveKeyboardHook();
        Unregister();
        Win32Properties.RemoveWndProcHookCallback(window, messageHook);
    }
}
