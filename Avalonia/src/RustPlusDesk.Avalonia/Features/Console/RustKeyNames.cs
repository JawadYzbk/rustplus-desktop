using Avalonia.Input;

namespace RustPlusDesk.Avalonia.Features.Console;

/// <summary>
/// Reference RustKeyNames: the name Rust's bind command expects for a key press (NumPad1 is "kp1", D1 is "1",
/// OemPlus is "="). Unknown keys return null rather than a guess, which would bind silently to nothing.
/// </summary>
public static class RustKeyNames
{
    /// <summary>Mouse buttons and the wheel never arrive as key presses, so they are offered as a list.</summary>
    public static readonly IReadOnlyList<string> MouseAndWheel = ["mouse1", "mouse2", "mouse3", "mouse4", "mouse5", "mousewheelup", "mousewheeldown"];

    private static readonly Dictionary<Key, string> Map = new()
    {
        [Key.NumPad0] = "kp0", [Key.NumPad1] = "kp1", [Key.NumPad2] = "kp2", [Key.NumPad3] = "kp3", [Key.NumPad4] = "kp4",
        [Key.NumPad5] = "kp5", [Key.NumPad6] = "kp6", [Key.NumPad7] = "kp7", [Key.NumPad8] = "kp8", [Key.NumPad9] = "kp9",
        [Key.Divide] = "kpdivide", [Key.Multiply] = "kpmultiply", [Key.Subtract] = "kpminus", [Key.Add] = "kpplus", [Key.Decimal] = "kpperiod",
        [Key.D0] = "0", [Key.D1] = "1", [Key.D2] = "2", [Key.D3] = "3", [Key.D4] = "4",
        [Key.D5] = "5", [Key.D6] = "6", [Key.D7] = "7", [Key.D8] = "8", [Key.D9] = "9",
        [Key.Space] = "space", [Key.Enter] = "return", [Key.Tab] = "tab", [Key.Back] = "backspace", [Key.Escape] = "escape",
        [Key.CapsLock] = "capslock", [Key.Insert] = "insert", [Key.Delete] = "delete", [Key.Home] = "home", [Key.End] = "end",
        [Key.PageUp] = "pageup", [Key.PageDown] = "pagedown", [Key.Up] = "uparrow", [Key.Down] = "downarrow",
        [Key.Left] = "leftarrow", [Key.Right] = "rightarrow",
        [Key.LeftShift] = "leftshift", [Key.RightShift] = "rightshift", [Key.LeftCtrl] = "leftcontrol", [Key.RightCtrl] = "rightcontrol",
        [Key.LeftAlt] = "leftalt", [Key.RightAlt] = "rightalt",
        [Key.OemPlus] = "=", [Key.OemMinus] = "-", [Key.OemComma] = ",", [Key.OemPeriod] = ".", [Key.OemQuestion] = "/",
        [Key.OemTilde] = "`", [Key.OemOpenBrackets] = "[", [Key.OemCloseBrackets] = "]", [Key.OemPipe] = "\\",
        [Key.OemSemicolon] = ";", [Key.OemQuotes] = "'",
    };

    /// <summary>Windows/menu keys never reach the game; binding them would do nothing.</summary>
    public static bool IsRejected(Key key) => key is Key.LWin or Key.RWin or Key.Apps or Key.None;

    public static string? ToRustName(Key key)
    {
        if (IsRejected(key)) return null;
        if (Map.TryGetValue(key, out var mapped)) return mapped;
        if (key is >= Key.A and <= Key.Z) return key.ToString().ToLowerInvariant();
        if (key is >= Key.F1 and <= Key.F24) return key.ToString().ToLowerInvariant();
        return null;
    }

    /// <summary>Reference ToRustBind: combinations as "[leftcontrol+leftshift+leftalt+key]", unless the key is a modifier.</summary>
    public static string? ToRustBind(Key key, KeyModifiers modifiers)
    {
        string? name = ToRustName(key);
        if (name is null) return null;
        bool modifierItself = key is Key.LeftShift or Key.RightShift or Key.LeftCtrl or Key.RightCtrl or Key.LeftAlt or Key.RightAlt;
        if (modifierItself) return name;
        var parts = new List<string>();
        if (modifiers.HasFlag(KeyModifiers.Control)) parts.Add("leftcontrol");
        if (modifiers.HasFlag(KeyModifiers.Shift)) parts.Add("leftshift");
        if (modifiers.HasFlag(KeyModifiers.Alt)) parts.Add("leftalt");
        if (parts.Count == 0) return name;
        parts.Add(name);
        return "[" + string.Join("+", parts) + "]";
    }
}
