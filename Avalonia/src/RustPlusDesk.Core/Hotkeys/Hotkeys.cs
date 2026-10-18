namespace RustPlusDesk.Core.Hotkeys;

/// <summary>Reference HotkeyOptions in hotkey_options.json: sequential with a delay, or parallel.</summary>
public sealed class HotkeyOptions
{
    public const int MinDelayMs = 50, MaxDelayMs = 1000, DelayStepMs = 50;
    public bool ParallelMode { get; set; }
    public int ToggleDelayMs { get; set; } = 150;
}

/// <summary>
/// The reference hotkeys.json: server key ("host:port", or "default") → gesture ("Ctrl+Alt+K") → entity IDs. One device
/// belongs to at most one gesture per server; several devices may share a gesture and toggle together.
/// </summary>
public sealed class HotkeyMap
{
    public const string DefaultServer = "default";
    public Dictionary<string, Dictionary<string, List<long>>> Servers { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Reference CurrentServerKey.</summary>
    public static string ServerKey(string? host, int? port) => string.IsNullOrEmpty(host) || port is null ? DefaultServer : $"{host}:{port}";

    public Dictionary<string, List<long>> ForServer(string server)
    {
        if (!Servers.TryGetValue(server, out var map)) Servers[server] = map = new(StringComparer.OrdinalIgnoreCase);
        return map;
    }

    public string? GestureFor(string server, long entityId) =>
        Servers.TryGetValue(server, out var map) ? map.FirstOrDefault(pair => pair.Value.Contains(entityId)).Key : null;

    /// <summary>Reference BtnSet: the device leaves every other gesture, then joins this one.</summary>
    public void Assign(string server, string gesture, long entityId)
    {
        var map = ForServer(server);
        foreach (var list in map.Values) list.Remove(entityId);
        if (!map.TryGetValue(gesture, out var ids)) map[gesture] = ids = [];
        if (!ids.Contains(entityId)) ids.Add(entityId);
    }

    public void Clear(string server, long entityId)
    {
        if (Servers.TryGetValue(server, out var map))
            foreach (var list in map.Values) list.Remove(entityId);
    }

    /// <summary>Removes deleted devices from one server's gestures (device delete cleanup).</summary>
    public bool Remove(string server, IReadOnlyCollection<long> entityIds)
    {
        if (!Servers.TryGetValue(server, out var map)) return false;
        bool changed = false;
        foreach (var list in map.Values) changed |= list.RemoveAll(entityIds.Contains) > 0;
        Prune();
        return changed;
    }

    /// <summary>Reference PruneEmptyGesturesAllServers: a gesture without devices would block the key for nothing.</summary>
    public void Prune()
    {
        foreach (var map in Servers.Values)
            foreach (string gesture in map.Where(pair => pair.Value is not { Count: > 0 }).Select(pair => pair.Key).ToList())
                map.Remove(gesture);
    }

    public HotkeyMap Clone()
    {
        var copy = new HotkeyMap();
        foreach (var (server, map) in Servers)
        {
            var target = copy.ForServer(server);
            foreach (var (gesture, ids) in map) target[gesture] = [.. ids];
        }
        return copy;
    }
}

/// <summary>Reference gesture text: "Ctrl+Alt+Shift+Win+Key" with the key's WPF/Avalonia name (A, D1, F5, NumPad1, OemPlus …).</summary>
public static class HotkeyGesture
{
    public static string Format(string keyName, bool ctrl, bool alt, bool shift, bool win)
    {
        var parts = new List<string>(5);
        if (ctrl) parts.Add("Ctrl");
        if (alt) parts.Add("Alt");
        if (shift) parts.Add("Shift");
        if (win) parts.Add("Win");
        parts.Add(keyName);
        return string.Join("+", parts);
    }

    /// <summary>Reference TryParseGesture: modifiers (Win or Meta) and exactly one key name.</summary>
    public static bool TryParse(string gesture, out bool ctrl, out bool alt, out bool shift, out bool win, out string key)
    {
        ctrl = alt = shift = win = false;
        key = "";
        foreach (string part in gesture.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (part.Equals("Ctrl", StringComparison.OrdinalIgnoreCase)) ctrl = true;
            else if (part.Equals("Alt", StringComparison.OrdinalIgnoreCase)) alt = true;
            else if (part.Equals("Shift", StringComparison.OrdinalIgnoreCase)) shift = true;
            else if (part.Equals("Win", StringComparison.OrdinalIgnoreCase) || part.Equals("Meta", StringComparison.OrdinalIgnoreCase)) win = true;
            else key = part;
        }
        return key.Length > 0;
    }
}

public interface IHotkeyStore
{
    Task<HotkeyMap> LoadMapAsync(CancellationToken ct);
    Task SaveMapAsync(HotkeyMap map, CancellationToken ct);
    Task<HotkeyOptions> LoadOptionsAsync(CancellationToken ct);
    Task SaveOptionsAsync(HotkeyOptions options, CancellationToken ct);
}

/// <summary>
/// System-wide hotkeys. Only some platforms allow an app to claim them; elsewhere <see cref="IsSupported"/> is false
/// and <see cref="UnavailableReason"/> says why, so the UI never pretends a binding is live.
/// </summary>
public interface IGlobalHotkeys
{
    bool IsSupported { get; }
    string UnavailableReason { get; }
    /// <summary>False when the gesture is invalid or another application holds it.</summary>
    bool Register(string gesture);
    void UnregisterAll();
    IReadOnlyDictionary<string, bool> RegistrationStatus { get; }
    event Action<string>? Pressed;
}
