using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Hotkeys;

namespace RustPlusDesk.Avalonia.Features.Hotkeys;

/// <summary>One editor row (reference HotkeysWindow.RowVM): a smart switch, or a saved ID whose device is gone.</summary>
public sealed partial class HotkeyRowViewModel(long entityId, string name, string status, bool isDeleted) : ObservableObject
{
    public long EntityId { get; } = entityId;
    public string DisplayName { get; } = name;
    public string Status { get; } = status;
    public bool IsDeleted { get; } = isDeleted;
    public string IdText => $"ID: {EntityId.ToString(CultureInfo.InvariantCulture)}";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasHotkey))]
    [NotifyPropertyChangedFor(nameof(HotkeyDisplay))]
    private string? hotkey;
    /// <summary>Reference warning: the last registration of this gesture failed (another app holds it).</summary>
    [ObservableProperty] private bool isRegisteredWarningVisible;
    public bool HasHotkey => !string.IsNullOrEmpty(Hotkey);
    public string HotkeyDisplay => HasHotkey ? Hotkey! : "—";
}

/// <summary>
/// The reference Hotkeys window: switches with their hotkey, Set (capture) and Clear, unknown saved IDs listed as
/// deleted, sequential/parallel mode with the 50–1000 ms toggle delay, and Close &amp; Activate / Close &amp; Deactivate.
/// It edits a copy that is applied only when the window closes.
/// </summary>
public sealed partial class HotkeysEditorViewModel : ObservableObject
{
    private readonly IReadOnlyDictionary<string, bool> registration;

    internal HotkeysEditorViewModel(HotkeyMap map, string server, HotkeyOptions options, IEnumerable<DeviceRowViewModel> devices,
        IReadOnlyDictionary<string, bool> registration, bool canActivate, string unavailableReason)
    {
        Map = map;
        Server = server;
        this.registration = registration;
        CanActivate = canActivate;
        UnavailableReason = unavailableReason;
        isParallel = options.ParallelMode;
        delayMs = Math.Clamp(options.ToggleDelayMs, HotkeyOptions.MinDelayMs, HotkeyOptions.MaxDelayMs);
        var switches = devices.Where(row => row.Definition.Kind == DeviceKind.Switch).DistinctBy(row => row.Definition.EntityId).ToList();
        var known = switches.Select(row => (long)row.Definition.EntityId).ToHashSet();
        var rows = switches.Select(row => new HotkeyRowViewModel(row.Definition.EntityId, row.Name,
            row.IsOn is bool on ? on ? "ON" : "OFF" : "–", false) { Hotkey = map.GestureFor(server, row.Definition.EntityId) }).ToList();
        foreach (var (gesture, ids) in map.ForServer(server))
            foreach (long id in ids.Where(id => !known.Contains(id) && rows.All(row => row.EntityId != id)))
                rows.Add(new HotkeyRowViewModel(id, "Unknown Device", "DELETED", true) { Hotkey = gesture });
        Rows = rows.OrderBy(row => row.DisplayName, StringComparer.CurrentCulture).ToList();
        foreach (var row in Rows) row.IsRegisteredWarningVisible = Warn(row.Hotkey);
    }

    public HotkeyMap Map { get; }
    public string Server { get; }
    public IReadOnlyList<HotkeyRowViewModel> Rows { get; }
    public bool HasRows => Rows.Count > 0;
    public bool CanActivate { get; }
    public string UnavailableReason { get; }
    public bool HasUnavailableReason => UnavailableReason.Length > 0;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSequential))]
    private bool isParallel;
    public bool IsSequential { get => !IsParallel; set => IsParallel = !value; }
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DelayText))]
    private int delayMs;
    public string DelayText => $"Toggle delay: {DelayMs.ToString(CultureInfo.CurrentCulture)} ms";
    /// <summary>Set when the window closes: true for Close &amp; Activate.</summary>
    public bool? Result { get; private set; }

    public HotkeyOptions Options => new() { ParallelMode = IsParallel, ToggleDelayMs = DelayMs };

    /// <summary>Reference slider: snaps to 50 ms steps.</summary>
    partial void OnDelayMsChanged(int value)
    {
        int snapped = Math.Clamp((int)Math.Round(value / (double)HotkeyOptions.DelayStepMs) * HotkeyOptions.DelayStepMs, HotkeyOptions.MinDelayMs, HotkeyOptions.MaxDelayMs);
        if (snapped != value) DelayMs = snapped;
    }

    public void Assign(HotkeyRowViewModel row, string gesture)
    {
        if (string.IsNullOrWhiteSpace(gesture)) return;
        Map.Assign(Server, gesture, row.EntityId);
        row.Hotkey = gesture;
        row.IsRegisteredWarningVisible = Warn(gesture);
    }

    [RelayCommand]
    private void Clear(HotkeyRowViewModel? row)
    {
        if (row is null) return;
        Map.Clear(Server, row.EntityId);
        row.Hotkey = null;
        row.IsRegisteredWarningVisible = false;
    }

    [RelayCommand(CanExecute = nameof(CanActivate))] private void CloseActivate() => Close(true);
    [RelayCommand] private void CloseDeactivate() => Close(false);

    public event EventHandler? CloseRequested;

    private void Close(bool activate)
    {
        Result = activate;
        CloseRequested?.Invoke(this, EventArgs.Empty);
    }

    private bool Warn(string? gesture) => !string.IsNullOrEmpty(gesture) && registration.TryGetValue(gesture, out bool ok) && !ok;
}

/// <summary>
/// Device hotkeys for the selected server (reference DEVICE HOTKEYS in MainWindow.xaml.cs): shared hotkeys.json and
/// hotkey_options.json, activation registers the server's gestures, a press toggles its switches once at a time with
/// a 350 ms per-gesture debounce, and switching servers re-registers that server's gestures while active.
/// </summary>
public sealed partial class HotkeysViewModel : ObservableObject
{
    private const int DebounceMs = 350;
    private readonly IHotkeyStore? store;
    private readonly Func<DateTimeOffset> clock;
    private readonly Dictionary<string, DateTimeOffset> lastPress = new(StringComparer.OrdinalIgnoreCase);
    private readonly SemaphoreSlim sequence = new(1, 1);
    private IGlobalHotkeys global = UnavailableGlobalHotkeys.ForCurrentPlatform();
    private Task pendingSave = Task.CompletedTask;

    public HotkeysViewModel(IHotkeyStore? store, Func<DateTimeOffset>? clock = null)
    {
        this.store = store;
        this.clock = clock ?? (() => DateTimeOffset.UtcNow);
    }

    public HotkeyMap Map { get; private set; } = new();
    public HotkeyOptions Options { get; private set; } = new();
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ButtonText))]
    private bool isActive;
    public string ButtonText => IsActive ? "Hotkeys active" : "Hotkeys";
    public bool IsSupported => global.IsSupported;
    public string UnavailableReason => global.UnavailableReason;
    public string ServerKey { get; private set; } = HotkeyMap.DefaultServer;

    /// <summary>Toggles the given entities (ids, parallel, delay) and reports how many confirmed.</summary>
    public Func<IReadOnlyCollection<uint>, bool, int, Task<int>>? Toggle { get; set; }
    public Action<string>? Log { get; set; }
    /// <summary>The shell opens the editor window for this model.</summary>
    public event EventHandler<HotkeysEditorViewModel>? EditRequested;
    /// <summary>Supplies the switches shown in the editor.</summary>
    public Func<IEnumerable<DeviceRowViewModel>>? Devices { get; set; }

    /// <summary>Installs the platform adapter once the main window exists (Win32 on Windows, otherwise unavailable).</summary>
    public void UseGlobalHotkeys(IGlobalHotkeys hotkeys)
    {
        global.Pressed -= OnPressed;
        global.UnregisterAll();
        global = hotkeys;
        global.Pressed += OnPressed;
        OnPropertyChanged(nameof(IsSupported));
        OnPropertyChanged(nameof(UnavailableReason));
        if (loaded) Activate();
    }
    private bool loaded;

    /// <summary>Reference OnSourceInitialized: load, then activate for the current server.</summary>
    public async Task LoadAsync(CancellationToken ct)
    {
        if (store is null) return;
        Map = await store.LoadMapAsync(ct);
        Options = await store.LoadOptionsAsync(ct);
        loaded = true;
        Activate();
    }

    public void SelectServer(string? host, int? port)
    {
        ServerKey = HotkeyMap.ServerKey(host, port);
        if (IsActive) Activate();
    }

    /// <summary>Reference ActivateHotkeysForCurrentServer: prune, re-register; active when any gesture registered.</summary>
    public void Activate()
    {
        Map.Prune();
        global.UnregisterAll();
        bool any = false;
        foreach (string gesture in Map.ForServer(ServerKey).Keys.ToList())
        {
            bool ok = global.Register(gesture);
            if (!ok && global.IsSupported) Log?.Invoke($"[hotkeys] Cannot register hotkey '{gesture}'.");
            any |= ok;
        }
        IsActive = any;
    }

    public void Deactivate()
    {
        global.UnregisterAll();
        IsActive = false;
    }

    /// <summary>Reference BtnHotkeys: deactivate, edit a copy, then save and activate or deactivate as chosen.</summary>
    [RelayCommand]
    private void Edit()
    {
        var status = new Dictionary<string, bool>(global.RegistrationStatus, StringComparer.OrdinalIgnoreCase);
        Deactivate();
        var editor = new HotkeysEditorViewModel(Map.Clone(), ServerKey, Options, Devices?.Invoke() ?? [], status, global.IsSupported, global.UnavailableReason);
        EditRequested?.Invoke(this, editor);
    }

    public void Apply(HotkeysEditorViewModel editor)
    {
        Map = editor.Map;
        Map.Prune();
        Options = editor.Options;
        QueueSave();
        if (editor.Result == true) Activate();
        else Deactivate();
    }

    /// <summary>Device delete cleanup: removes the deleted IDs from this server's gestures and re-registers.</summary>
    public void RemoveDevices(string? host, int? port, IReadOnlyCollection<long> entityIds)
    {
        if (!Map.Remove(HotkeyMap.ServerKey(host, port), entityIds)) return;
        QueueSave();
        if (IsActive) Activate();
    }

    internal void OnPressed(string gesture)
    {
        var now = clock();
        if (lastPress.TryGetValue(gesture, out var last) && (now - last).TotalMilliseconds < DebounceMs) return;
        lastPress[gesture] = now;
        if (!Map.ForServer(ServerKey).TryGetValue(gesture, out var ids) || ids.Count == 0) return;
        _ = RunAsync(ids.Where(id => id is > 0 and <= uint.MaxValue).Select(id => (uint)id).Distinct().ToList());
    }

    /// <summary>Reference RunHotkeySequenceOnceAsync: a second press while a sequence runs is ignored.</summary>
    private async Task RunAsync(IReadOnlyCollection<uint> ids)
    {
        if (Toggle is null) return;
        if (!await sequence.WaitAsync(0)) { Log?.Invoke("[hotkeys] Hotkey sequence already running – ignored."); return; }
        try { await Toggle(ids, Options.ParallelMode, Options.ToggleDelayMs); }
        catch (Exception) { Log?.Invoke("[hotkeys] The hotkey toggle could not be completed."); }
        finally { sequence.Release(); }
    }

    private void QueueSave()
    {
        if (store is null) return;
        var map = Map.Clone();
        var options = new HotkeyOptions { ParallelMode = Options.ParallelMode, ToggleDelayMs = Options.ToggleDelayMs };
        pendingSave = SaveAfterAsync(pendingSave, map, options);
    }

    private async Task SaveAfterAsync(Task previous, HotkeyMap map, HotkeyOptions options)
    {
        await previous;
        try
        {
            await store!.SaveMapAsync(map, CancellationToken.None);
            await store.SaveOptionsAsync(options, CancellationToken.None);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Log?.Invoke("[hotkeys] Hotkeys could not be saved. Check file permissions and disk space."); }
    }

    public Task FlushAsync() => pendingSave;

    /// <summary>Waits for a running toggle sequence (tests and shutdown).</summary>
    public async Task WaitIdleAsync()
    {
        await sequence.WaitAsync();
        sequence.Release();
    }
}
