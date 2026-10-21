using System.ComponentModel;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Avalonia.Features.Connections;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MainWindow's side of the Widget Overlay: it owns the data, the dock only draws.
///
/// Every member here is a read of state that already exists in the connection - the dock is a second view of the server's
/// status, not a second source of truth. The parts the connection does not hold - the logic rules, the oil rig alarms,
/// the session and the death markers - are answered by the context the shell hands over.
/// </summary>
public sealed class ConnectionDockHost : IDockHost, IDockHostDevices, IDockHostContext, IDisposable
{
    private readonly ConnectionViewModel connection;
    private readonly IDockHostContext? context;
    private readonly ChatViewModel chat;
    private readonly List<DockChatLine> teamLines = [];
    private readonly List<DockChatLine> clanLines = [];

    /// <summary>Reference MainWindow's DockHost: the dock reads the connection's state once a second.</summary>
    public ConnectionDockHost(ConnectionViewModel connection, IDockHostContext? context = null)
    {
        this.connection = connection;
        this.context = context;
        chat = connection.Chat;
        chat.LiveMessage += OnLiveMessage;
        chat.PropertyChanged += OnChatChanged;
    }

    /// <summary>Reference DockServerTime: the server's clock, as "HH:mm".</summary>
    public string ServerTime => connection.ServerTime;

    /// <summary>Reference DockTimeUntilNextPhase: the time until the next sunrise or sunset, or null.</summary>
    public string? TimeUntilNextPhase =>
        string.IsNullOrWhiteSpace(connection.NextPhase) ? null : connection.NextPhase;

    /// <summary>Reference DockPopulation: the players on the server, as "128 / 200".</summary>
    public string Population => connection.Population;

    /// <summary>Reference DockPopulation's queue: the players waiting for a slot, or nothing when there are none.</summary>
    public string Queue => connection.Queue is "-" or "" ? "" : connection.Queue;

    // ---- the devices -------------------------------------------------------------------------

    /// <summary>
    /// Reference ICommandDockHost.DockDevices: the paired devices of the connected server.
    ///
    /// The list is read again every time the dock asks, because a device's state changes while the dock is open and a
    /// tile that showed a stale list would offer to switch something that has moved on.
    /// </summary>
    public IReadOnlyList<DockHostDevice> DockDevices => connection.Devices.Devices
        .SelectMany(row => row.DescendantsAndSelf())
        .Where(row => row.Definition.Kind != DeviceKind.Group && row.Definition.EntityId != 0)
        .GroupBy(row => row.Definition.EntityId)
        .Select(group => group.First())
        .Select(row => new DockHostDevice(row.Definition.EntityId, row.Name, KindOf(row.Definition.Kind))
        {
            IsOn = row.IsOn,
            IsMissing = row.IsMissing,
            UpkeepSeconds = UpkeepSeconds(row.Storage),
            IconId = row.Definition.CustomIconId,
            IconShortName = row.Definition.CustomIconShortName,
        })
        .ToArray();

    /// <summary>
    /// Reference ICommandDockHost.ToggleDockSwitchAsync: a smart switch is flipped through the same guarded path the
    /// device list uses, so a tile cannot drive a switch the panel would refuse.
    /// </summary>
    public async Task<bool> SetDockDeviceAsync(uint entityId, bool on, CancellationToken ct) =>
        await connection.Devices.SetSwitchStateAsync(entityId, on, ct).ConfigureAwait(false);

    /// <summary>Reference the reference's device kind names, which the dock's picture and the picker's glyph are chosen from.</summary>
    private static string KindOf(DeviceKind kind) => kind switch
    {
        DeviceKind.Switch => "SmartSwitch",
        DeviceKind.Alarm => "SmartAlarm",
        DeviceKind.StorageMonitor => "StorageMonitor",
        DeviceKind.Group => "Group",
        _ => "Unknown",
    };

    /// <summary>Reference SmartDevice.UpkeepSeconds (protect seconds left); unknown when the device has no protection.</summary>
    private static int? UpkeepSeconds(DeviceStorage? storage) =>
        storage?.HasProtection == true && storage.ProtectionExpiry is { } expiry
            ? Math.Max(0, (int)(expiry - DateTimeOffset.UtcNow).TotalSeconds)
            : null;

    // ---- the rest of the app's state ---------------------------------------------------------

    /// <summary>Reference ICommandDockHost.DockRules: the active server's Logic Engine rules, enabled or not.</summary>
    public IReadOnlyList<DockHostRule> DockRules => context?.DockRules ?? [];

    /// <summary>Reference ICommandDockHost.IsDockLogicEngineActive: false while the Logic Engine's master switch is off.</summary>
    public bool IsDockLogicEngineActive => context?.IsDockLogicEngineActive == true;

    /// <summary>Reference ICommandDockHost.RunDockRule: a rule is run as if the Widget Overlay trigger had fired.</summary>
    public void RunDockRule(string ruleId) => context?.RunDockRule(ruleId);

    /// <summary>Reference DockOilRigAlarms: the newest oil rig alarm of the active server.</summary>
    public DockHostEvent? LatestOilRigAlarm => context?.LatestOilRigAlarm;

    /// <summary>Reference ICommandDockHost.DockServerKey: the active server as "{host}-{port}".</summary>
    public string? DockServerKey => context?.DockServerKey;

    /// <summary>Reference SessionTracker.Instance.For: the session of the active server.</summary>
    public SessionState? DockSession => context?.DockSession;

    /// <summary>Reference ICommandDockHost.DockDeathMarkerCount: how many death markers the map is showing.</summary>
    public int DockDeathMarkerCount => context?.DockDeathMarkerCount ?? 0;

    /// <summary>Reference ICommandDockHost.WipeDockDeathMarkers: the map's death markers are cleared.</summary>
    public Task WipeDockDeathMarkersAsync(bool keepLatest, CancellationToken ct) =>
        context?.WipeDockDeathMarkersAsync(keepLatest, ct) ?? Task.CompletedTask;

    /// <summary>Reference MainWindow.Camera: the map the workspace draws, which the dock's map tile draws as well.</summary>
    public MapViewModel? Map => context?.Map;

    /// <summary>Reference the workspace's death markers: the pins the map is showing.</summary>
    public IReadOnlyList<DeathPin>? DeathPins => context?.DeathPins;

    /// <summary>Reference MapShowDeathTags: whether the death markers are drawn on the map.</summary>
    public bool ShowDeathPins => context?.ShowDeathPins == true;

    /// <summary>Reference ShowMiniMap: the map is brought forward.</summary>
    public void OpenMap() => context?.OpenMap();

    // ---- the chat ----------------------------------------------------------------------------

    /// <summary>
    /// Reference GetDockChat: the last lines of the team chat, or of the clan chat.
    ///
    /// The chat keeps one channel's lines on screen at a time, so the adapter keeps a copy of both: the channel that is on
    /// screen is read straight from the chat, and the other one is kept up to date with the messages that arrive live.
    /// </summary>
    public IReadOnlyList<DockChatLine> Chat(bool clan, int count)
    {
        SyncChannel();
        var feed = clan ? clanLines : teamLines;
        return feed.Count <= count ? feed.ToList() : feed.GetRange(feed.Count - count, count);
    }

    /// <summary>Reference the live chat: a message that arrived while the dock was watching.</summary>
    private void OnLiveMessage(ChatMessage message)
    {
        var line = ToDockLine(message);
        (message.Channel == ChatChannel.Clan ? clanLines : teamLines).Add(line);
    }

    /// <summary>Reference the chat's history being read on connect: both channels are taken then.</summary>
    private void OnChatChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(ChatViewModel.IsConnected)) return;

        if (!chat.IsConnected)
        {
            teamLines.Clear();
            clanLines.Clear();
            return;
        }

        Seed();
    }

    /// <summary>Reads the channel that is on screen into the adapter's copy of it.</summary>
    private void SyncChannel()
    {
        // A search filters the list on screen, so it is left alone: the dock keeps the whole feed.
        if (chat.IsSearchOpen) return;

        var target = chat.IsClan ? clanLines : teamLines;
        target.Clear();
        foreach (var line in chat.Lines) target.Add(ToDockLine(line));
    }

    /// <summary>
    /// Reads both channels' history once, when the chat connects.
    ///
    /// The chat shows one channel at a time, so the other one is read by looking at it for a moment and then putting the
    /// chat back where it was.
    /// </summary>
    private void Seed()
    {
        bool wasClan = chat.IsClan;
        try
        {
            chat.IsClan = false;
            teamLines.Clear();
            foreach (var line in chat.Lines) teamLines.Add(ToDockLine(line));

            chat.IsClan = true;
            clanLines.Clear();
            foreach (var line in chat.Lines) clanLines.Add(ToDockLine(line));
        }
        finally
        {
            chat.IsClan = wasClan;
        }
    }

    /// <summary>Reference CommandDockChatLine: one line of the chat the dock shows.</summary>
    private static DockChatLine ToDockLine(ChatLine line) =>
        new(line.Name, line.Text, line.Message.Time.LocalDateTime);

    /// <summary>Reference CommandDockChatLine: one line of the chat the dock shows.</summary>
    private static DockChatLine ToDockLine(ChatMessage message) =>
        new(string.IsNullOrWhiteSpace(message.Name) ? message.SteamId.ToString() : message.Name,
            message.Text, message.Time.LocalDateTime);

    /// <summary>Reference the dock being closed: the chat is left as it was found.</summary>
    public void Dispose()
    {
        chat.LiveMessage -= OnLiveMessage;
        chat.PropertyChanged -= OnChatChanged;
    }
}
