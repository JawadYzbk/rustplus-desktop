using System.Globalization;
using RustPlusDesk.Core.Alerts;

namespace RustPlusDesk.Core.Map;

/// <summary>Which cargo announcements are on (reference AnnounceCargo/CargoDocking/CargoEgress/CargoArrival, master applied).</summary>
public sealed record CargoAlertFlags(bool Spawn, bool Docking, bool Egress, bool Arrival, bool Master);

/// <summary>
/// What the reference learns per server host in tracking_settings.json: the docking duration (LearnedDockingDurations,
/// default 8 min, 1–60), the arrival trigger point per harbor (ServerCargoTriggers) and a full cargo run
/// (LearnedCargoFullLifeMinutes, 10–120).
/// </summary>
public interface ICargoMemory
{
    int DockingMinutes(string host);
    void SetDockingMinutes(string host, int minutes);
    (double X, double Y)? TriggerPoint(string host, string harbor);
    void SetTriggerPoint(string host, string harbor, double x, double y);
    bool HasAnyTrigger(string host);
    int FullLifeMinutes(string host);
    void SetFullLife(string host, int minutes);
}

/// <summary>One live cargo ship's dock state, as the reference event dock reads _cargoDockStates.</summary>
public sealed record CargoShipView(ulong Id, double LastX, double LastY, DateTime? FirstSeen, DateTime? DockTime, bool IsDocked,
    bool WasAlreadyDocked, bool SeenAtEdge, string? HarborName);

/// <summary>
/// Reference ProcessCargoDocking + CleanupCargoDockStates for live cargo markers: a cargo first seen after the first poll is
/// announced ("far out at sea (near {grid})" beyond 105 % of the half world), docking at a harbor (stationary under 0.5 m
/// within 300 m) is announced after 5 s, an arrival warning fires near a learned trigger point while approaching, and a
/// departure warning 5 minutes before the learned docking time ends. Docking durations, trigger points (5 minutes before
/// docking) and full runs are learned; a ship unseen for 60 s is gone.
/// </summary>
public sealed class CargoTracker(ICargoMemory memory)
{
    public const string FarOutAtSea = "far out at sea (near {0})";
    private readonly Dictionary<ulong, State> ships = [];
    private bool firstPoll = true;

    public AlertTemplates Templates { get; set; } = AlertTemplates.Defaults;

    /// <summary>Reference connect reset (_firstPollDyn = true, states cleared).</summary>
    public void Reset()
    {
        ships.Clear();
        firstPoll = true;
    }

    /// <summary>When the last tracked ship was despawned (reference _cargoLastDespawnUtc), for the event dock.</summary>
    public DateTime? LastDespawnUtc { get; private set; }

    /// <summary>The live ships, as the event dock reads the reference _cargoDockStates.</summary>
    public IReadOnlyList<CargoShipView> Ships => ships.Select(entry => new CargoShipView(entry.Key, entry.Value.LastX, entry.Value.LastY,
        entry.Value.FirstSeen, entry.Value.DockTime, entry.Value.IsDocked, entry.Value.WasAlreadyDocked, entry.Value.SeenAtEdge, entry.Value.HarborName)).ToArray();

    /// <summary>One marker poll: the chat messages to send and the log lines to write.</summary>
    public (IReadOnlyList<string> Messages, IReadOnlyList<string> Log) Observe(IReadOnlyList<MapMarker> markers, double worldSize,
        IReadOnlyList<MapMonument> monuments, string host, CargoAlertFlags flags, DateTime now)
    {
        var messages = new List<string>();
        var log = new List<string>();
        var harbors = monuments.Where(monument => monument.Token.Contains("Harbor", StringComparison.OrdinalIgnoreCase)).ToArray();
        foreach (var cargo in markers.Where(marker => marker.Kind == MapMarkerKind.CargoShip))
            Process(cargo, worldSize, harbors, host, flags, now, messages, log);
        firstPoll = false;
        Cleanup(host, now, log);
        return (messages, log);
    }

    private void Process(MapMarker m, double worldSize, MapMonument[] harbors, string host, CargoAlertFlags flags, DateTime now, List<string> messages, List<string> log)
    {
        if (!ships.TryGetValue(m.Id, out var state))
        {
            state = new State { LastX = m.X, LastY = m.Y, FirstSeen = now };
            if (!firstPoll)
            {
                // Cargo spawns at the outer edge: more than 85 % of the half world from the centre.
                double half = worldSize * 0.5, cx = m.X - half, cy = m.Y - half;
                double fromCentre = Math.Sqrt(cx * cx + cy * cy);
                state.SeenAtEdge = fromCentre > half * 0.85;
                if (flags.Spawn)
                {
                    string grid = MapGrid.Label(worldSize, m.X, m.Y);
                    string where = fromCentre > half * 1.05 ? string.Format(CultureInfo.InvariantCulture, FarOutAtSea, grid) : grid;
                    messages.Add(Templates.Format("AlertCargoSpawned", where));
                    if (state.SeenAtEdge)
                        log.Add(string.Create(CultureInfo.InvariantCulture, $"[cargo] Spawn detected at edge (dist: {fromCentre:F0}, threshold: {half * 0.85:F0})"));
                }
            }
            ships[m.Id] = state;
        }
        state.LastSeen = now;
        state.History.Add((now, m.X, m.Y));
        if (state.History.Count > 150) state.History.RemoveAll(entry => entry.At < now.AddMinutes(-12));

        double dx = m.X - state.LastX, dy = m.Y - state.LastY;
        double moved = Math.Sqrt(dx * dx + dy * dy);
        bool stationary = moved < 0.5;
        if (!stationary) state.EverMoved = true;

        if (stationary && !state.IsDocked)
        {
            var harbor = harbors.FirstOrDefault(h => Math.Sqrt(Math.Pow(h.X - m.X, 2) + Math.Pow(h.Y - m.Y, 2)) < 300);
            if (harbor is not null)
            {
                state.IsDocked = true;
                state.DockTime = now;
                state.HarborName = MonumentCatalog.DisplayName(harbor.Token);
                state.AnnouncedDock = false;
                state.AnnouncedEgress = false;
                state.HarborCount++;
                // Stationary the very first time it was seen: it was already docked, no dock alert.
                if (!state.EverMoved && state.LastX == m.X && state.LastY == m.Y)
                {
                    state.WasAlreadyDocked = true;
                    state.AnnouncedDock = true;
                }
                // Learn the arrival trigger point: where it was about 5 minutes ago.
                if (!state.WasAlreadyDocked)
                {
                    var target = now.AddMinutes(-5);
                    var best = state.History.OrderBy(entry => Math.Abs((entry.At - target).TotalSeconds)).FirstOrDefault();
                    if (best.At != default && (now - best.At).TotalMinutes > 4) memory.SetTriggerPoint(host, harbor.Token, best.X, best.Y);
                }
            }
        }
        else if (moved > 2.0 && state.IsDocked)
        {
            if (state.DockTime is { } docked && !state.WasAlreadyDocked && (now - docked).TotalMinutes > 2)
                memory.SetDockingMinutes(host, (int)Math.Round((now - docked).TotalMinutes));
            state.IsDocked = false;
            state.DockTime = null;
            state.AnnouncedDock = false;
            state.AnnouncedEgress = false;
            state.WasAlreadyDocked = false;
            state.AnnouncedArrival = false;
        }

        if (state.IsDocked && !state.AnnouncedDock && flags.Docking && state.DockTime is { } dockTime && (now - dockTime).TotalSeconds >= 5)
        {
            messages.Add(Templates.Format("AlertCargoDocked", state.HarborName ?? "", MapGrid.Label(worldSize, m.X, m.Y)));
            state.AnnouncedDock = true;
        }

        if (!state.IsDocked && !state.AnnouncedArrival && flags.Arrival)
        {
            foreach (var h in harbors)
            {
                if (memory.TriggerPoint(host, h.Token) is not { } point) continue;
                if (Math.Sqrt(Math.Pow(m.X - point.X, 2) + Math.Pow(m.Y - point.Y, 2)) >= 150) continue;
                double toHarbor = Math.Sqrt(Math.Pow(m.X - h.X, 2) + Math.Pow(m.Y - h.Y, 2));
                double lastToHarbor = Math.Sqrt(Math.Pow(state.LastX - h.X, 2) + Math.Pow(state.LastY - h.Y, 2));
                if (toHarbor >= lastToHarbor) continue;
                messages.Add(Templates.Format("AlertCargoExpectedDock", MonumentCatalog.DisplayName(h.Token), MapGrid.Label(worldSize, h.X, h.Y)));
                state.AnnouncedArrival = true;
                state.ArrivalWarnedAt = now;
                break;
            }
        }

        if (state.IsDocked && state.DockTime is { } since && !state.AnnouncedEgress && flags.Master)
        {
            int duration = memory.DockingMinutes(host);
            double elapsed = (now - since).TotalMinutes;
            if (elapsed >= duration - 5 && duration > 5)
            {
                if (!flags.Egress)
                    log.Add(string.Create(CultureInfo.InvariantCulture, $"[Egress] BLOCKED: AnnounceCargoEgress=False (duration={duration}m, elapsed={elapsed:F1}m)"));
                else messages.Add(Templates.Format("AlertCargoDeparting", state.HarborName ?? "", MapGrid.Label(worldSize, m.X, m.Y)));
                state.AnnouncedEgress = true;
            }
        }

        state.LastX = m.X;
        state.LastY = m.Y;
    }

    /// <summary>Reference CleanupCargoDockStates: 60 s unseen = despawned; learn a full run and check the arrival warning.</summary>
    private void Cleanup(string host, DateTime now, List<string> log)
    {
        foreach (var (id, state) in ships.ToArray())
        {
            if ((now - state.LastSeen).TotalSeconds <= 60) continue;
            LastDespawnUtc = now;
            log.Add(string.Create(CultureInfo.InvariantCulture, $"[cargo] Despawn detected – last seen {(now - state.LastSeen).TotalSeconds:F0}s ago."));
            if (state.FirstSeen is { } first && state.HarborCount >= 1 && state.SeenAtEdge)
            {
                int total = (int)(now - first).TotalMinutes;
                if (total > 20) memory.SetFullLife(host, total);
            }
            if (state.ArrivalWarnedAt is { } warned && state.DockTime is { } docked && state.HarborName is { } harbor)
            {
                double warnToDock = (docked - warned).TotalMinutes;
                if (warnToDock < 2.0 || warnToDock > 8.0)
                {
                    log.Add(string.Create(CultureInfo.InvariantCulture, $"[cargo] Trigger point for {harbor} discarded (warn-to-dock: {warnToDock:F1}m, expected ~5m). Will re-learn on next run."));
                    memory.SetTriggerPoint(host, harbor, 0, 0);
                }
                else log.Add(string.Create(CultureInfo.InvariantCulture, $"[cargo] Trigger point for {harbor} confirmed (warn-to-dock: {warnToDock:F1}m)."));
            }
            ships.Remove(id);
        }
    }

    private sealed class State
    {
        public double LastX, LastY;
        public DateTime? DockTime, FirstSeen, ArrivalWarnedAt;
        public DateTime LastSeen;
        public bool AnnouncedDock, AnnouncedEgress, AnnouncedArrival, IsDocked, WasAlreadyDocked, SeenAtEdge, EverMoved;
        public string? HarborName;
        public int HarborCount;
        public List<(DateTime At, double X, double Y)> History { get; } = [];
    }
}

/// <summary>
/// <see cref="ICargoMemory"/> over the shared tracking_settings.json, with the reference property names and limits
/// (TrackingService.Get/SetLearnedDockingDuration, Get/SetCargoTriggerPoint, HasAnyCargoTrigger, SetLearnedCargoFullLife).
/// </summary>
public sealed class SettingsCargoMemory(Settings.IAppSettingsStore settings) : ICargoMemory
{
    public const string DockingKey = "LearnedDockingDurations", TriggersKey = "ServerCargoTriggers", FullLifeKey = "LearnedCargoFullLifeMinutes";
    public const int DefaultDockingMinutes = 8;

    public int DockingMinutes(string host) =>
        settings.Get<Dictionary<string, int>?>(DockingKey, null) is { } all && all.TryGetValue(host, out int minutes) ? minutes : DefaultDockingMinutes;

    public void SetDockingMinutes(string host, int minutes)
    {
        if (minutes is < 1 or > 60) return;
        var all = settings.Get<Dictionary<string, int>?>(DockingKey, null) ?? [];
        all[host] = minutes;
        settings.Set(DockingKey, all);
    }

    public (double X, double Y)? TriggerPoint(string host, string harbor) =>
        Triggers().TryGetValue(host, out var points) && points.TryGetValue(harbor, out var point) ? (point.X, point.Y) : null;

    public void SetTriggerPoint(string host, string harbor, double x, double y)
    {
        var all = Triggers();
        if (!all.TryGetValue(host, out var points)) all[host] = points = [];
        points[harbor] = new TriggerEntry(x, y);
        settings.Set(TriggersKey, all);
    }

    public bool HasAnyTrigger(string host) => Triggers().TryGetValue(host, out var points) && points.Count > 0;

    /// <summary>Reference TrackingService.GetLearnedCargoFullLife: 0 when the server never taught a full run.</summary>
    public int FullLifeMinutes(string host) =>
        settings.Get<Dictionary<string, int>?>(FullLifeKey, null) is { } all && all.TryGetValue(host, out int minutes) ? minutes : 0;

    public void SetFullLife(string host, int minutes)
    {
        if (minutes is < 10 or > 120) return;
        var all = settings.Get<Dictionary<string, int>?>(FullLifeKey, null) ?? [];
        all[host] = minutes;
        settings.Set(FullLifeKey, all);
    }

    private Dictionary<string, Dictionary<string, TriggerEntry>> Triggers() =>
        settings.Get<Dictionary<string, Dictionary<string, TriggerEntry>>?>(TriggersKey, null) ?? [];

    /// <summary>Reference CargoTriggerPoint {X, Y}.</summary>
    public sealed record TriggerEntry(double X, double Y);
}
