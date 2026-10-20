using System.Text.Json.Serialization;
using RustPlusDesk.Core.Time;

namespace RustPlusDesk.Core.Dock;

/// <summary>
/// Reference SessionSnapshot: what the tracker needs to know about right now, sampled once a second.
///
/// Everything it counts comes from here - the connection, the server, the player's own position and idleness, the
/// population of the server and the team with each member's state - so the tracker itself holds no game knowledge at all.
/// </summary>
public sealed record SessionSnapshot(
    bool Connected,
    string? ServerKey,
    ulong MySteamId,
    double? MyX,
    double? MyY,
    bool MyAfk,
    int Population,
    IReadOnlyList<(ulong SteamId, bool IsDead)> Team);

/// <summary>
/// Reference PopulationSample: one reading of how many players are on the server.
///
/// The name in the saved file is the reference's own, "Utc", so both apps read each other's file; the reading is a
/// moment in time and the count that went with it.
/// </summary>
public sealed record PopulationSample(
    [property: JsonPropertyName("Utc")] DateTime At,
    int Players);

/// <summary>
/// Reference SessionState: one play session on one server.
///
/// A session is a local thing - it starts when you sit down and ends when you stop - and nothing about it needs the
/// cloud to be right. The property names are the reference's, because the file is the reference's file.
/// </summary>
public sealed class SessionState
{
    /// <summary>Reference SessionState.ServerKey: the server this session belongs to, as "{host}-{port}".</summary>
    public string ServerKey { get; set; } = "";

    /// <summary>Reference SessionState.StartedUtc: when the session began, in UTC.</summary>
    public DateTime StartedUtc { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Reference SessionState.LastSeenUtc: when the tracker last saw this session.
    ///
    /// A restart that lands inside <see cref="SessionTracker.ResumeWindow"/> carries on rather than starting over -
    /// closing the app to fix a setting should not cost you the evening's numbers.
    /// </summary>
    public DateTime LastSeenUtc { get; set; } = DateTime.UtcNow;

    /// <summary>Reference SessionState.OnlineSeconds: how long the player has been on this server.</summary>
    public double OnlineSeconds { get; set; }

    /// <summary>Reference SessionState.AfkSeconds: how much of that time the player was away from the keyboard.</summary>
    public double AfkSeconds { get; set; }

    /// <summary>Reference SessionState.DistanceMetres: metres walked, in world units, which Rust measures in metres.</summary>
    public double DistanceMetres { get; set; }

    /// <summary>Reference SessionState.Deaths: how many times the player has died this session.</summary>
    public int Deaths { get; set; }

    /// <summary>Reference SessionState.TeamDeaths: how many times somebody else on the team has died this session.</summary>
    public int TeamDeaths { get; set; }

    /// <summary>
    /// Reference SessionState.Population: the population readings, one every five minutes, that the server-info tile's
    /// population line is drawn from.
    /// </summary>
    public List<PopulationSample> Population { get; set; } = new();
}

/// <summary>
/// Reference DataManager.SaveCache/LoadCache for the session: the session of one server, kept under its own key.
///
/// The tracker keeps its sessions here and nowhere else, so the numbers survive a restart without the tracker having to
/// know where they are written.
/// </summary>
public interface ISessionTrackerStore
{
    /// <summary>The session saved for the server, or null when there is none or it cannot be read.</summary>
    SessionState? Load(string serverKey);

    /// <summary>The session is written under the server's own key, replacing what was there.</summary>
    void Save(SessionState state);
}

/// <summary>
/// Reference DispatcherTimer: the heartbeat the tracker samples on, once a second.
///
/// It is an interface so the tests can fire a tick without waiting for one, and so the host can decide which thread the
/// sampling happens on.
/// </summary>
public interface ISessionTimer
{
    /// <summary>How often the tick fires.</summary>
    TimeSpan Interval { get; set; }

    /// <summary>Starts ticking.</summary>
    void Start();

    /// <summary>Stops ticking.</summary>
    void Stop();

    /// <summary>Raised on every tick.</summary>
    event EventHandler? Ticking;
}

/// <summary>
/// The real one-second heartbeat, on a background thread.
///
/// The reference runs its timer on the UI thread; here the sampling is plain work on a background one, because the
/// tracker only reads numbers and writes a file, and neither belongs on the thread that draws the dock.
/// </summary>
public sealed class SystemSessionTimer : ISessionTimer, IDisposable
{
    private readonly object gate = new();
    private System.Threading.Timer? timer;
    private TimeSpan interval = TimeSpan.FromSeconds(1);

    /// <summary>How often the tick fires. Changing it while the timer runs takes effect at once.</summary>
    public TimeSpan Interval
    {
        get { lock (gate) return interval; }
        set
        {
            lock (gate)
            {
                interval = value;
                timer?.Change(value, value);
            }
        }
    }

    /// <summary>Raised on every tick.</summary>
    public event EventHandler? Ticking;

    /// <summary>Starts ticking.</summary>
    public void Start()
    {
        lock (gate)
        {
            timer ??= new System.Threading.Timer(OnTick, null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            timer.Change(interval, interval);
        }
    }

    /// <summary>Stops ticking.</summary>
    public void Stop()
    {
        lock (gate) timer?.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
    }

    /// <summary>Stops the timer and lets it go.</summary>
    public void Dispose()
    {
        lock (gate)
        {
            timer?.Dispose();
            timer = null;
        }
    }

    private void OnTick(object? state) => Ticking?.Invoke(this, EventArgs.Empty);
}

/// <summary>
/// Reference SessionTracker: keeps a running session per server - time online, ground covered, time idle, deaths.
///
/// It samples rather than subscribes. Half of what it counts is elapsed time, so a timer has to run either way, and
/// pulling one snapshot a second is both simpler and impossible to leak across the connect and disconnect cycles the
/// alternative would have to survive.
///
/// Stored locally, per server, so the WPF app and this one read each other's numbers.
/// </summary>
public sealed class SessionTracker : IDisposable
{
    /// <summary>A gap longer than this starts a new session instead of continuing the old.</summary>
    public static readonly TimeSpan ResumeWindow = TimeSpan.FromMinutes(30);

    /// <summary>
    /// How often the population is written down.
    ///
    /// Five minutes rather than an hour: the reading is already in hand from the status the app polls anyway, so a
    /// longer gap buys nothing and costs the graph - a three-hour evening would be three dots. At this rate the same
    /// evening is a line.
    /// </summary>
    public static readonly TimeSpan PopulationInterval = TimeSpan.FromMinutes(5);

    /// <summary>Movement below this is noise in the position feed, not walking.</summary>
    private const double MinStepMetres = 0.5;

    /// <summary>Above this, the player teleported, respawned or the map changed.</summary>
    private const double MaxStepMetres = 500;

    /// <summary>Half a day of readings, which is longer than any tile draws.</summary>
    private const int MaxPopulationSamples = 144;

    /// <summary>
    /// How long one tick may count for.
    ///
    /// A laptop that was asleep, or an app that was busy, must not add an hour of "online" to the evening's numbers.
    /// </summary>
    private const double MaxTickSeconds = 5;

    /// <summary>
    /// How often the session is written to disk.
    ///
    /// The numbers move every second, but the only thing that reads them off disk is the next app start, and a file
    /// rewritten sixty times a minute for that is a waste of a disk.
    /// </summary>
    private static readonly TimeSpan SaveInterval = TimeSpan.FromSeconds(30);

    private readonly object gate = new();
    private readonly IClock clock;
    private readonly ISessionTimer timer;
    private readonly ISessionTrackerStore store;
    private readonly Dictionary<string, SessionState> sessions = new();
    private readonly Dictionary<ulong, bool> deadLastTick = new();

    private double? lastX;
    private double? lastY;
    private string? scratchServerKey;
    private string? trackedServerKey;
    private DateTime lastTickUtc;
    private DateTime lastSaveUtc = DateTime.MinValue;
    private bool disposed;

    /// <summary>
    /// Reference SessionTracker.Instance: the tracker, with the clock it reads the time from, the heartbeat it samples on
    /// and the file it keeps its sessions in.
    /// </summary>
    public SessionTracker(IClock clock, ISessionTimer timer, ISessionTrackerStore store)
    {
        this.clock = clock;
        this.timer = timer;
        this.store = store;
        lastTickUtc = UtcNow;
        timer.Interval = TimeSpan.FromSeconds(1);
        timer.Ticking += OnTick;
        timer.Start();
    }

    /// <summary>Set once by the main window. Null until then, and the tracker idles.</summary>
    public Func<SessionSnapshot?>? Source { get; set; }

    /// <summary>The session for a server, loaded from disk if this one is not it.</summary>
    public SessionState? For(string? serverKey)
    {
        if (string.IsNullOrEmpty(serverKey)) return null;

        lock (gate)
        {
            if (sessions.TryGetValue(serverKey, out SessionState? live) && live is not null) return live;
        }

        return store.Load(serverKey);
    }

    /// <summary>Throws the session away and starts counting again from now.</summary>
    public void Wipe(string? serverKey)
    {
        if (string.IsNullOrEmpty(serverKey)) return;

        DateTime now = UtcNow;
        var fresh = new SessionState { ServerKey = serverKey, StartedUtc = now, LastSeenUtc = now };

        lock (gate)
        {
            sessions[serverKey] = fresh;
            lastX = lastY = null;
            deadLastTick.Clear();
        }

        store.Save(fresh);
    }

    /// <summary>
    /// Reference SessionTracker.Tick: one sample of the world, and the numbers it moves.
    ///
    /// Public so a test can drive the tracker a tick at a time, and so a host that would rather tick by hand can.
    /// </summary>
    public void Tick()
    {
        DateTime now = UtcNow;
        double elapsed;
        lock (gate)
        {
            elapsed = Math.Clamp((now - lastTickUtc).TotalSeconds, 0, MaxTickSeconds);
            lastTickUtc = now;
        }

        SessionSnapshot? snapshot;
        try { snapshot = Source?.Invoke(); }
        catch { return; }

        if (snapshot == null || !snapshot.Connected || string.IsNullOrEmpty(snapshot.ServerKey))
        {
            // Not counting, and not resetting either: the resume window decides whether
            // coming back continues this session or starts the next one.
            Persist(force: true);
            return;
        }

        SessionState session = EnsureSession(snapshot.ServerKey, now);

        session.OnlineSeconds += elapsed;
        if (snapshot.MyAfk) session.AfkSeconds += elapsed;
        session.LastSeenUtc = now;

        AccumulateDistance(session, snapshot);
        CountDeaths(session, snapshot);
        SamplePopulation(session, snapshot);

        Persist(force: false);
    }

    /// <summary>
    /// Reference SessionTracker.SamplePopulation: writes down the population, at most one reading every five minutes.
    /// </summary>
    public void SamplePopulation(SessionState session, SessionSnapshot snapshot)
    {
        if (snapshot.Population <= 0) return;

        DateTime now = UtcNow;
        PopulationSample? last = session.Population.Count > 0 ? session.Population[^1] : null;
        if (last != null && now - last.At < PopulationInterval) return;

        session.Population.Add(new PopulationSample(now, snapshot.Population));

        if (session.Population.Count > MaxPopulationSamples) session.Population.RemoveAt(0);
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Session.FormatSpan: a duration in the narrowest form that is still unambiguous.
    ///
    /// The cell is a fifth of the tile, so "10h 15m" was being cut off by the column beside it. Past an hour it becomes a
    /// clock reading - 10:15h - which drops a unit label and a space without losing anything, since the glyph above it is
    /// already a clock. Below an hour there is room for the label, and a bare "15" under a clock would read as a time of day.
    /// </summary>
    public static string FormatSpan(TimeSpan span)
    {
        if (span.TotalHours >= 1) return $"{(int)span.TotalHours}:{span.Minutes:D2}h";
        if (span.TotalMinutes >= 1) return $"{span.Minutes}m";
        return $"{span.Seconds}s";
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Session.FormatDistance: kilometres past a thousand metres, and never a decimal
    /// below it.
    ///
    /// The distance is a sum of straight lines between position samples, so it is a floor - a figure with a decimal point
    /// would claim a precision the sampling does not have.
    /// </summary>
    public static string FormatDistance(double metres)
        => metres >= 1000 ? $"{metres / 1000.0:0.0} km" : $"{(int)metres} m";

    /// <summary>Stops the heartbeat.</summary>
    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        timer.Ticking -= OnTick;
        timer.Stop();
    }

    private DateTime UtcNow => clock.Now.UtcDateTime;

    private void OnTick(object? sender, EventArgs e) => Tick();

    private SessionState EnsureSession(string serverKey, DateTime now)
    {
        lock (gate)
        {
            // Another server is another world: the last position and the last known states belong to the one just left.
            if (!string.Equals(scratchServerKey, serverKey, StringComparison.Ordinal))
            {
                scratchServerKey = serverKey;
                lastX = lastY = null;
                deadLastTick.Clear();
            }

            if (sessions.TryGetValue(serverKey, out SessionState? live) && live is not null
                && now - live.LastSeenUtc <= ResumeWindow)
            {
                trackedServerKey = serverKey;
                return live;
            }

            SessionState? stored = store.Load(serverKey);

            bool resumable = stored != null
                             && stored.ServerKey == serverKey
                             && now - stored.LastSeenUtc <= ResumeWindow;

            SessionState session = resumable
                ? stored!
                : new SessionState { ServerKey = serverKey, StartedUtc = now, LastSeenUtc = now };

            // A resumed session knows nothing about where the player was while the app was shut,
            // so the first position after it comes back is a starting point, not a step.
            sessions[serverKey] = session;
            trackedServerKey = serverKey;
            lastX = lastY = null;
            deadLastTick.Clear();

            return session;
        }
    }

    private void AccumulateDistance(SessionState session, SessionSnapshot snapshot)
    {
        if (snapshot.MyX is not { } x || snapshot.MyY is not { } y)
        {
            lastX = lastY = null;
            return;
        }

        if (lastX is { } px && lastY is { } py)
        {
            double step = Math.Sqrt((x - px) * (x - px) + (y - py) * (y - py));

            // Between two samples the player walked a curve and this counts the chord, so
            // the total is a floor rather than an exact figure - which is why the tile
            // rounds it instead of claiming decimals it does not have.
            if (step is > MinStepMetres and < MaxStepMetres) session.DistanceMetres += step;
        }

        lastX = x;
        lastY = y;
    }

    private void CountDeaths(SessionState session, SessionSnapshot snapshot)
    {
        foreach (var (steamId, isDead) in snapshot.Team)
        {
            bool wasDead = deadLastTick.TryGetValue(steamId, out bool previous) && previous;
            deadLastTick[steamId] = isDead;

            // The transition, not the state: a body lies there until the player respawns,
            // and counting the state would add a death every second it did.
            if (!isDead || wasDead) continue;

            if (steamId == snapshot.MySteamId) session.Deaths++;
            else session.TeamDeaths++;
        }
    }

    private void Persist(bool force)
    {
        SessionState? session;
        lock (gate)
        {
            if (trackedServerKey is null) return;
            if (!sessions.TryGetValue(trackedServerKey, out session) || session is null) return;

            DateTime now = UtcNow;
            if (!force && now - lastSaveUtc < SaveInterval) return;
            lastSaveUtc = now;
        }

        try { store.Save(session); }
        catch { /* a session is not worth an exception on the UI thread */ }
    }
}
