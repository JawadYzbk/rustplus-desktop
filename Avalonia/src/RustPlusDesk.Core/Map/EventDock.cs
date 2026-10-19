namespace RustPlusDesk.Core.Map;

/// <summary>
/// One row of the reference event dock (MainWindow.EventDockItem): the active flag, the marker to centre on, whether a
/// click can follow it, and the timer line with its tooltip.
/// </summary>
public sealed record EventDockItem(string Key, string Name, string Icon, bool Active, ulong Id, double X, double Y,
    bool Trackable, string? TimerText, string? Tip);

/// <summary>
/// Reference event dock (MainWindow.Map.Markers.UpdateEventDock / RefreshEventDock): the "events" panel over the map with
/// a row for each patrol helicopter, cargo ship, CH47, travelling vendor, deep-sea event and oil-rig crate timer. A row
/// stays visible for one poll after its marker disappears (GetPersistentEvent keeps it while MissingCount &lt; 5) and the
/// event bookkeeping that feeds the timers happens when the marker is finally dropped (MissingCount reaches 10): a shot
/// down helicopter records "was inside the map", a vendor records its despawn time, and cargo falls back to the last
/// despawn time. Spawn times come from the first poll a new key was seen; a marker already present on the first poll of a
/// session is "connected mid-event", so no spawn time is known. A connect resets the session state (the deep-sea times
/// included), exactly like the reference _firstPollDyn handling.
/// </summary>
public sealed class EventDockModel
{
    /// <summary>Reference GetPersistentEvent: a row is still shown for four polls after the marker went missing.</summary>
    public const int DisplayMissingPolls = 5;
    /// <summary>Reference removal bookkeeping happens at MissingCount == 10.</summary>
    public const int RemoveMissingPolls = 10;

    /// <summary>Reference HeliEventName / CargoShip / Chinook / Vendor / DeepSea resources.</summary>
    public const string HeliName = "Patrol Heli", CargoName = "Cargo Ship", ChinookName = "Chinook (CH47)",
        VendorName = "Travelling Vendor", DeepSeaName = "Deep Sea Event";

    /// <summary>Reference resource files under Assets/icons (and Assets/Markers in the port).</summary>
    public const string HeliIcon = "patrol_helicopter.png", CargoIcon = "cargo.png", ChinookIcon = "ch47.png",
        VendorIcon = "vendor.png", DeepSeaIcon = "ds_event.png", CrateIcon = "crate.png";

    private sealed class KindState
    {
        public string? Key;
        public MapMarker? Last;
        public int Missing;
    }

    private readonly HashSet<string> known = new(StringComparer.Ordinal);
    private KindState heli = new(), chinook = new(), cargo = new(), vendor = new();
    private bool firstPoll = true;
    /// <summary>The world size of the last poll, for the grid label of the chat "heli" answer.</summary>
    private double lastWorldSize;

    private DateTime? heliSpawnTime, heliLastEventUtc;
    private bool heliMidEvent, heliLastEventWasCrash;
    private DateTime? vendorSpawnTime, vendorDespawnTime;
    private bool vendorMidEvent;
    private bool deepSeaActive, deepSeaMidEvent;
    private DateTime? deepSeaSpawnTime, deepSeaDespawnTime;

    /// <summary>
    /// Reference connect reset: _firstPollDyn = true and the live dock states cleared. The reference keeps the last
    /// session's helicopter/vendor/cargo bookkeeping fields, so a dead event still shows "… ago this session" after a
    /// reconnect; the deep-sea spawn/despawn times are cleared with the connection.
    /// </summary>
    public void Reset()
    {
        firstPoll = true;
        known.Clear();
        heli = new KindState();
        chinook = new KindState();
        cargo = new KindState();
        vendor = new KindState();
        deepSeaActive = false;
        deepSeaMidEvent = false;
        deepSeaSpawnTime = null;
        deepSeaDespawnTime = null;
    }

    /// <summary>
    /// One marker poll, in the reference event order (helicopter, cargo, CH47, vendor, deep sea, oil rigs). Timers:
    /// <paramref name="cargoLastDespawnUtc"/> and the cargo views come from <see cref="CargoTracker"/>,
    /// <paramref name="deepSeaChange"/> from <see cref="DeepSeaTracker"/> and the rig rows from <see cref="OilRigTimer"/>
    /// plus a recent-alarm lookup.
    /// </summary>
    public IReadOnlyList<EventDockItem> Observe(IReadOnlyList<MapMarker> markers, double worldSize,
        IReadOnlyList<CargoShipView> cargoShips, DateTime? cargoLastDespawnUtc, int cargoFullLifeMinutes, int cargoDockingMinutes,
        DeepSeaChange deepSeaChange, DateTime now, OilRigTimer? rigs = null, Func<string, DateTime?>? rigAlarmAt = null)
    {
        ApplyDeepSeaChange(deepSeaChange, now);
        lastWorldSize = worldSize;

        ProcessKind(heli, Marker(markers, MapMarkerKind.PatrolHelicopter), worldSize, now);
        ProcessKind(cargo, Marker(markers, MapMarkerKind.CargoShip), worldSize, now);
        ProcessKind(chinook, Marker(markers, MapMarkerKind.Chinook), worldSize, now);
        ProcessKind(vendor, Marker(markers, MapMarkerKind.TravellingVendor), worldSize, now);

        var items = new List<EventDockItem>(6);
        AddHeli(items, now);
        AddCargo(items, now, cargoShips, cargoLastDespawnUtc, cargoFullLifeMinutes, cargoDockingMinutes);
        AddChinook(items);
        AddVendor(items, now);
        AddDeepSea(items, now);
        AddOilRigs(items, now, rigs, rigAlarmAt);
        firstPoll = false;
        return items;
    }

    /// <summary>
    /// The raw event state the reference chat commands read off MainWindow (_dynStates, _heliSpawnTime,
    /// _heliLastEventUtc, _vendorSpawnTime, _vendorDespawnTime, _deepSeaSpawnTime, ...). "Active" is the marker
    /// still being tracked (MissingCount below <see cref="RemoveMissingPolls"/>), which is what the reference's
    /// dynamic-state dictionary holds, not the shorter dock-row display window.
    /// </summary>
    public Chat.ChatEventCommandState EventState() => new(
        DeepSeaActive: deepSeaActive,
        DeepSeaSpawnTime: deepSeaSpawnTime,
        DeepSeaDespawnTime: deepSeaDespawnTime,
        HeliActive: heli.Last is not null,
        HeliSpawnTime: heliSpawnTime,
        HeliGrid: heli.Last is { } marker ? MapGrid.Label(lastWorldSize, marker.X, marker.Y) : null,
        HeliLastEventUtc: heliLastEventUtc,
        HeliLastEventWasCrash: heliLastEventWasCrash,
        VendorActive: vendor.Last is not null,
        VendorSpawnTime: vendorSpawnTime,
        VendorDespawnTime: vendorDespawnTime);

    private void ApplyDeepSeaChange(DeepSeaChange change, DateTime now)
    {
        switch (change)
        {
            case DeepSeaChange.Spawned:
                deepSeaActive = true;
                deepSeaSpawnTime = now;
                deepSeaDespawnTime = null;
                deepSeaMidEvent = false;
                break;
            case DeepSeaChange.ActiveOnFirstPoll:
                deepSeaActive = true;
                deepSeaSpawnTime = null;
                deepSeaMidEvent = true;
                break;
            case DeepSeaChange.Despawned:
                deepSeaActive = false;
                deepSeaDespawnTime = now;
                deepSeaMidEvent = false;
                break;
            case DeepSeaChange.InactiveOnFirstPoll:
                deepSeaActive = false;
                deepSeaDespawnTime = null;
                deepSeaMidEvent = false;
                break;
        }
    }

    private static MapMarker? Marker(IReadOnlyList<MapMarker> markers, MapMarkerKind kind) =>
        markers.FirstOrDefault(marker => marker.Kind == kind);

    private void ProcessKind(KindState state, MapMarker? marker, double worldSize, DateTime now)
    {
        if (marker is { } present)
        {
            // Reference: a key not in _dynKnown is a new event of that kind; a known key is the same event refreshed.
            string key = Key(present);
            if (known.Add(key)) Learned(state, present, now);
            state.Key = key;
            state.Last = present;
            state.Missing = 0;
            return;
        }
        if (state.Last is null) return;
        state.Missing++;
        if (state.Missing < RemoveMissingPolls) return;
        Gone(state.Last, worldSize, now);
        state.Key = null;
        state.Last = null;
        state.Missing = 0;
    }

    /// <summary>Reference marker key: the entity id when there is one, else a rounded-position fallback.</summary>
    private static string Key(MapMarker marker) => marker.Id != 0
        ? $"{marker.Kind}:{marker.Id}"
        : $"{marker.Kind}:{marker.X / 10:F0}:{marker.Y / 10:F0}";

    private static MapMarker? Display(KindState state) =>
        state.Last is { } marker && state.Missing < DisplayMissingPolls ? marker : null;

    private void Learned(KindState state, MapMarker marker, DateTime now)
    {
        if (marker.Kind == MapMarkerKind.PatrolHelicopter)
        {
            heliMidEvent = firstPoll;
            heliSpawnTime = firstPoll ? null : now;
        }
        else if (marker.Kind == MapMarkerKind.TravellingVendor)
        {
            vendorMidEvent = firstPoll;
            vendorSpawnTime = firstPoll ? null : now;
            vendorDespawnTime = null;
        }
    }

    private void Gone(MapMarker marker, double worldSize, DateTime now)
    {
        if (marker.Kind == MapMarkerKind.PatrolHelicopter)
        {
            heliLastEventUtc = now;
            // Reference: a shot-down helicopter is one whose last position was still inside the map.
            heliLastEventWasCrash = worldSize > 0 && marker.X > 0 && marker.X < worldSize && marker.Y > 0 && marker.Y < worldSize;
            heliSpawnTime = null;
            heliMidEvent = false;
        }
        else if (marker.Kind == MapMarkerKind.TravellingVendor)
        {
            vendorDespawnTime = now;
            vendorSpawnTime = null;
            vendorMidEvent = false;
        }
    }

    private void AddHeli(List<EventDockItem> items, DateTime now)
    {
        var marker = Display(heli);
        bool active = marker is not null;
        string? timer = null, tip = null;
        if (marker is not null)
        {
            if (heliSpawnTime is { } spawn)
            {
                var running = now - spawn;
                timer = Elapsed(running);
                tip = $"Patrol Heli active - running for {Ago(running)}";
            }
            else
            {
                timer = "??:??";
                tip = heliMidEvent ? "Connected mid-event – spawn time unknown" : "Patrol Heli active (spawn time unknown)";
            }
        }
        else if (heliLastEventUtc is { } last)
        {
            var ago = now - last;
            timer = Since(ago);
            tip = heliLastEventWasCrash
                ? $"Patrol Heli shot down {Ago(ago)} ago this session"
                : $"Patrol Heli left the map {Ago(ago)} ago this session";
        }
        items.Add(new EventDockItem("heli", HeliName, HeliIcon, active, marker?.Id ?? 0, marker?.X ?? 0, marker?.Y ?? 0, true, timer, tip));
    }

    private void AddCargo(List<EventDockItem> items, DateTime now, IReadOnlyList<CargoShipView> ships,
        DateTime? lastDespawnUtc, int fullLifeMinutes, int dockingMinutes)
    {
        var marker = Display(cargo);
        bool active = marker is not null;
        string? timer = null, tip = null;
        if (marker is not null)
        {
            var ship = marker.Id != 0 ? ships.FirstOrDefault(candidate => candidate.Id == marker.Id) : null;
            if (ship is { IsDocked: true, DockTime: { } dockTime })
            {
                if (dockingMinutes > 0 && !ship.WasAlreadyDocked)
                {
                    var remaining = TimeSpan.FromMinutes(dockingMinutes) - (now - dockTime);
                    timer = remaining.TotalSeconds > 0 ? Countdown(remaining) : "0:00";
                    tip = $"Docked at {ship.HarborName ?? "harbor"} - departs in ~{timer}";
                }
                else
                {
                    timer = "??:??";
                    tip = ship.WasAlreadyDocked
                        ? "Already docked on connect – depart time unknown"
                        : $"Docked at {ship.HarborName ?? "harbor"} (dock duration not yet learned)";
                }
            }
            else if (ship is { SeenAtEdge: true, FirstSeen: { } firstSeen } && fullLifeMinutes > 0)
            {
                var remaining = TimeSpan.FromMinutes(fullLifeMinutes) - (now - firstSeen);
                if (remaining.TotalSeconds > 0)
                {
                    timer = Countdown(remaining);
                    tip = "Time remaining on map (from spawn)";
                }
            }
            else if (ship is { SeenAtEdge: false })
            {
                tip = fullLifeMinutes > 0
                    ? $"Connected mid-route – time unknown (route ~{fullLifeMinutes}m total)"
                    : "Connected mid-route – time unknown";
            }
        }
        else if (lastDespawnUtc is { } despawn)
        {
            var ago = now - despawn;
            timer = Since(ago);
            tip = $"Cargo Ship despawned {(int)ago.TotalMinutes}m {ago.Seconds}s ago this session";
        }
        items.Add(new EventDockItem("cargo", CargoName, CargoIcon, active, marker?.Id ?? 0, marker?.X ?? 0, marker?.Y ?? 0, true, timer, tip));
    }

    private void AddChinook(List<EventDockItem> items)
    {
        var marker = Display(chinook);
        items.Add(new EventDockItem("chinook", ChinookName, ChinookIcon, marker is not null, marker?.Id ?? 0, marker?.X ?? 0,
            marker?.Y ?? 0, true, null, null));
    }

    private void AddVendor(List<EventDockItem> items, DateTime now)
    {
        var marker = Display(vendor);
        bool active = marker is not null;
        string? timer = null, tip = null;
        if (marker is not null)
        {
            if (vendorSpawnTime is { } spawn)
            {
                var running = now - spawn;
                timer = Elapsed(running);
                tip = $"Travelling Vendor active - running for {Ago(running)}";
            }
            else
            {
                timer = "??:??";
                tip = vendorMidEvent ? "Connected mid-event – spawn time unknown" : "Travelling Vendor active (spawn time unknown)";
            }
        }
        else if (vendorDespawnTime is { } despawn)
        {
            var ago = now - despawn;
            timer = Since(ago);
            tip = $"Travelling Vendor left the map {Ago(ago)} ago this session";
        }
        items.Add(new EventDockItem("vendor", VendorName, VendorIcon, active, marker?.Id ?? 0, marker?.X ?? 0, marker?.Y ?? 0, true, timer, tip));
    }

    private void AddDeepSea(List<EventDockItem> items, DateTime now)
    {
        string? timer = null, tip = null;
        if (deepSeaActive)
        {
            if (deepSeaSpawnTime is { } spawn)
            {
                var running = now - spawn;
                timer = Elapsed(running);
                tip = $"Deep Sea active - running for {Ago(running)}";
            }
            else
            {
                timer = "??:??";
                tip = deepSeaMidEvent ? "Connected mid-event – spawn time unknown" : "Deep Sea active (spawn time unknown)";
            }
        }
        else if (deepSeaDespawnTime is { } despawn)
        {
            var ago = now - despawn;
            timer = Since(ago);
            tip = $"Deep Sea Event ended {Ago(ago)} ago this session";
        }
        items.Add(new EventDockItem("deepsea", DeepSeaName, DeepSeaIcon, deepSeaActive, 0, 0, 0, false, timer, tip));
    }

    /// <summary>
    /// Reference BuildOilRigAlarmDockItems + CreateOilRigDockItem: a row per rig that has a running crate timer or a
    /// recorded "Oil Rig" smart alarm, otherwise one generic row. The reference keys the rows off the Logic Engine's
    /// StartTimer rules and shows the alarm's local timestamp; the port has neither, so it keys the rows off the live
    /// timer and shows the countdown while one runs (PARITY-PLAN.md records the deviation). The rig rows are not
    /// trackable: the port has no per-rig crate marker to centre on, so a click is ignored.
    /// </summary>
    private static void AddOilRigs(List<EventDockItem> items, DateTime now, OilRigTimer? rigs, Func<string, DateTime?>? rigAlarmAt)
    {
        if (rigs is null && rigAlarmAt is null) return;
        bool smallRig = Shown(OilRigTimer.SmallRig);
        bool largeRig = Shown(OilRigTimer.LargeRig);
        if (smallRig) Add(OilRigTimer.SmallRig, "Small Oil", "oilrig_small");
        if (largeRig) Add(OilRigTimer.LargeRig, "Large Oil", "oilrig_large");
        if (!smallRig && !largeRig) Add(OilRigTimer.GenericRig, "Oil Rig Smart Alarm", "oilrig");
        return;

        // The reference shows a rig row whenever that rig has a rule or a recorded alarm.
        bool Shown(string rig) => rigs?.TimeLeft(rig, now) > TimeSpan.Zero || rigAlarmAt?.Invoke(rig) is not null;

        void Add(string rig, string name, string key)
        {
            var left = rigs?.TimeLeft(rig, now);
            DateTime? alarm = rigAlarmAt?.Invoke(rig);
            bool timerRunning = left is { } value && value > TimeSpan.Zero;
            bool alarmRecent = alarm is { } at && now - at >= TimeSpan.Zero && now - at < TimeSpan.FromMinutes(15);
            string text = alarm is { } stamp ? stamp.ToLocalTime().ToString("HH:mm:ss")
                : timerRunning ? OilRigTimer.FormatDuration(left!.Value) : "-";
            string tip = timerRunning
                ? $"{rig}: crate unlocks in ~{OilRigTimer.FormatDuration(left!.Value)}"
                : alarmRecent ? $"{rig}: alarm at {alarm!.Value.ToLocalTime():HH:mm:ss}"
                : $"{rig}: no crate timer running";
            items.Add(new EventDockItem(key, name, CrateIcon, timerRunning || alarmRecent, 0, 0, 0, false, text, tip));
        }
    }

    private static string Elapsed(TimeSpan span) => span.TotalHours >= 1
        ? $"{(int)span.TotalHours}h {span.Minutes:D2}m"
        : $"{(int)span.TotalMinutes}m {span.Seconds:D2}s";

    private static string Since(TimeSpan span) => $"-{(int)span.TotalMinutes}:{span.Seconds:D2}";

    private static string Ago(TimeSpan span) => span.TotalHours >= 1
        ? $"{(int)span.TotalHours}h {span.Minutes}m"
        : $"{(int)span.TotalMinutes}m";

    private static string Countdown(TimeSpan span) => $"{(int)span.TotalMinutes}:{span.Seconds:D2}";
}
