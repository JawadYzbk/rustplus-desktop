using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Core.Time;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference SessionTracker and MiniMapWindow.CommandDock.Session: what this sitting has amounted to on this server.
///
/// The tracker counts the time online, the ground covered, the time idle, the deaths and the team's deaths, and keeps
/// the population readings the server-info tile's line is drawn from. The figures are written the way the reference writes
/// them, because somebody who knows the old app should not have to learn a new dialect.
/// </summary>
[TestClass]
public sealed class SessionTrackerTests
{
    private const string ServerA = "10.0.0.1-28082";
    private const string ServerB = "10.0.0.2-28082";
    private const ulong MySteamId = 76561198000000001;
    private const ulong FriendSteamId = 76561198000000002;

    private static readonly DateTimeOffset Start = new(2026, 10, 20, 12, 0, 0, TimeSpan.Zero);

    private sealed class FakeTimer : ISessionTimer
    {
        public TimeSpan Interval { get; set; } = TimeSpan.FromSeconds(1);

        public bool Running { get; private set; }

        public event EventHandler? Ticking;

        public void Start() => Running = true;

        public void Stop() => Running = false;

        /// <summary>One tick, the way the real timer raises one.</summary>
        public void Fire() => Ticking?.Invoke(this, EventArgs.Empty);
    }

    private sealed class MemoryStore : ISessionTrackerStore
    {
        public Dictionary<string, SessionState> Saved { get; } = new(StringComparer.Ordinal);

        public SessionState? Load(string serverKey)
            => Saved.TryGetValue(serverKey, out SessionState? state) ? state : null;

        public void Save(SessionState state) => Saved[state.ServerKey] = state;
    }

    private sealed class Folder : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "rustplus-sessions-" + Guid.NewGuid().ToString("N"));

        public Folder() => Directory.CreateDirectory(Path);

        public void Dispose()
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
        }
    }

    private sealed class Harness : IDisposable
    {
        public FixedClock Clock { get; }

        public FakeTimer Timer { get; } = new();

        public SessionTracker Tracker { get; }

        public Harness(ISessionTrackerStore store, DateTimeOffset? start = null)
        {
            Clock = new FixedClock(start ?? Start);
            Tracker = new SessionTracker(Clock, Timer, store);
        }

        /// <summary>Moves the clock on and fires a tick, the way the real timer does once a second.</summary>
        public void Tick(double seconds = 1)
        {
            Clock.Now = Clock.Now.AddSeconds(seconds);
            Timer.Fire();
        }

        public void Dispose() => Tracker.Dispose();
    }

    private static SessionSnapshot Snapshot(
        string serverKey = ServerA,
        ulong mySteamId = MySteamId,
        double? x = null,
        double? y = null,
        bool afk = false,
        int population = 0,
        IReadOnlyList<(ulong SteamId, bool IsDead)>? team = null)
        => new(true, serverKey, mySteamId, x, y, afk, population, team ?? [(mySteamId, false)]);

    // ---- the numbers ---------------------------------------------------------------------------

    [TestMethod]
    public void TheTimeOnlineIsCounted()
    {
        using var tracker = new Harness(new MemoryStore());
        tracker.Tracker.Source = () => Snapshot();

        tracker.Tick();
        tracker.Tick();

        SessionState? session = tracker.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(2, session.OnlineSeconds, 0.001, "Two ticks a second apart are two seconds online.");
    }

    [TestMethod]
    public void TheDistanceWalkedGrowsWhenThePlayerMovesAndNotWhenTheyDoNot()
    {
        using var tracker = new Harness(new MemoryStore());
        double x = 0;
        double y = 0;
        tracker.Tracker.Source = () => Snapshot(x: x, y: y);

        tracker.Tick();          // the first position is a starting point, not a step
        x = 10;
        tracker.Tick();          // ten metres walked
        tracker.Tick();          // standing still

        SessionState? session = tracker.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(10, session.DistanceMetres, 0.001, "Only the step that was walked is counted.");
    }

    [TestMethod]
    public void TheIdleTimeIsCounted()
    {
        using var tracker = new Harness(new MemoryStore());
        bool afk = false;
        tracker.Tracker.Source = () => Snapshot(afk: afk);

        tracker.Tick();
        afk = true;
        tracker.Tick();
        tracker.Tick();
        afk = false;
        tracker.Tick();

        SessionState? session = tracker.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(2, session.AfkSeconds, 0.001, "Only the time away from the keyboard is idle time.");
        Assert.AreEqual(4, session.OnlineSeconds, 0.001, "Idle time is still time online.");
    }

    [TestMethod]
    public void TheDeathsAreCountedWhenThePlayerDiesAndNotWhileTheBodyLiesThere()
    {
        using var tracker = new Harness(new MemoryStore());
        bool dead = false;
        tracker.Tracker.Source = () => Snapshot(team: [(MySteamId, dead)]);

        tracker.Tick();          // alive
        dead = true;
        tracker.Tick();          // the death
        tracker.Tick();          // still lying there
        dead = false;
        tracker.Tick();          // respawned
        dead = true;
        tracker.Tick();          // died again

        SessionState? session = tracker.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(2, session.Deaths, "A death is the transition, not the state.");
        Assert.AreEqual(0, session.TeamDeaths, "The player's own death is not a team death.");
    }

    [TestMethod]
    public void TheTeamsDeathsAreCounted()
    {
        using var tracker = new Harness(new MemoryStore());
        bool friendDead = false;
        tracker.Tracker.Source = () => Snapshot(team: [(MySteamId, false), (FriendSteamId, friendDead)]);

        tracker.Tick();
        friendDead = true;
        tracker.Tick();
        tracker.Tick();

        SessionState? session = tracker.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(1, session.TeamDeaths, "One teammate died once.");
        Assert.AreEqual(0, session.Deaths, "Somebody else's death is not the player's own.");
    }

    // ---- the population ------------------------------------------------------------------------

    [TestMethod]
    public void ThePopulationIsSampledEveryFiveMinutes()
    {
        using var tracker = new Harness(new MemoryStore());
        tracker.Tracker.Source = () => Snapshot(population: 42);

        tracker.Tick();              // the first reading
        tracker.Tick(60);            // a minute later: still the same reading
        tracker.Tick(4 * 60);        // five minutes after the first: a new one

        SessionState? session = tracker.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(2, session.Population.Count, "One reading every five minutes, not one every second.");
        Assert.AreEqual(42, session.Population[0].Players);
        Assert.AreEqual(42, session.Population[1].Players);
    }

    [TestMethod]
    public void AServerWithNobodyOnItIsNotSampled()
    {
        using var tracker = new Harness(new MemoryStore());
        tracker.Tracker.Source = () => Snapshot(population: 0);

        tracker.Tick();
        tracker.Tick(10 * 60);

        SessionState? session = tracker.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(0, session.Population.Count, "A population of nobody is not a reading.");
    }

    // ---- the restart --------------------------------------------------------------------------

    [TestMethod]
    public void ASessionSurvivesARestartInsideTheResumeWindow()
    {
        using var folder = new Folder();
        var store = new LegacySessionTrackerStore(folder.Path);

        using (var first = new Harness(store))
        {
            first.Tracker.Source = () => Snapshot();
            first.Tick();
            first.Tick();
            first.Tracker.Source = null;     // the app is closed
            first.Tick();                    // which saves the session
        }

        using var second = new Harness(store, Start.AddMinutes(10));
        second.Tracker.Source = () => Snapshot();
        second.Tick();

        SessionState? session = second.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(3, session.OnlineSeconds, 0.001,
            "The two seconds from before the restart, and the one since it.");
        Assert.AreEqual(Start.AddSeconds(1).UtcDateTime, session.StartedUtc,
            "The session did not start again: it began on the first tick before the restart.");
    }

    [TestMethod]
    public void ASessionDoesNotSurviveARestartPastTheResumeWindow()
    {
        using var folder = new Folder();
        var store = new LegacySessionTrackerStore(folder.Path);

        using (var first = new Harness(store))
        {
            first.Tracker.Source = () => Snapshot();
            first.Tick();
            first.Tracker.Source = null;
            first.Tick();
        }

        using var second = new Harness(store, Start.AddMinutes(31));
        second.Tracker.Source = () => Snapshot();
        second.Tick();

        SessionState? session = second.Tracker.For(ServerA);
        Assert.IsNotNull(session);
        Assert.AreEqual(1, session.OnlineSeconds, 0.001, "Past the window it is a new session, counting only from now.");
        Assert.AreEqual(Start.AddMinutes(31).AddSeconds(1).UtcDateTime, session.StartedUtc,
            "Past the window the session starts again.");
    }

    // ---- one server at a time -----------------------------------------------------------------

    [TestMethod]
    public void TheSessionsOfDifferentServersAreSeparate()
    {
        using var tracker = new Harness(new MemoryStore());
        string serverKey = ServerA;
        tracker.Tracker.Source = () => Snapshot(serverKey: serverKey);

        tracker.Tick();
        serverKey = ServerB;
        tracker.Tick();
        tracker.Tick();

        SessionState? first = tracker.Tracker.For(ServerA);
        SessionState? second = tracker.Tracker.For(ServerB);
        Assert.IsNotNull(first);
        Assert.IsNotNull(second);
        Assert.AreNotSame(first, second, "Each server has its own session.");
        Assert.AreEqual(1, first.OnlineSeconds, 0.001, "The first server was counted only while it was the one being played.");
        Assert.AreEqual(2, second.OnlineSeconds, 0.001, "The second server has been counted on its own.");

        serverKey = ServerA;
        tracker.Tick();
        Assert.AreEqual(2, first.OnlineSeconds, 0.001, "Coming back to the first server continues its session.");
        Assert.AreEqual(2, second.OnlineSeconds, 0.001, "The second server was not counted while another was being played.");
    }

    // ---- the figures --------------------------------------------------------------------------

    [TestMethod]
    public void TheDurationIsWrittenTheWayTheReferenceWritesIt()
    {
        CultureInfo culture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

            Assert.AreEqual("1:15h", SessionTracker.FormatSpan(TimeSpan.FromMinutes(75)),
                "Past an hour it is a clock reading, without the unit label.");
            Assert.AreEqual("10:15h", SessionTracker.FormatSpan(TimeSpan.FromMinutes(615)),
                "Three figures of hours are still a clock reading.");
            Assert.AreEqual("15m", SessionTracker.FormatSpan(TimeSpan.FromMinutes(15)),
                "Below an hour there is room for the unit.");
            Assert.AreEqual("35s", SessionTracker.FormatSpan(TimeSpan.FromSeconds(35)),
                "Below a minute it is seconds.");
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
        }
    }

    [TestMethod]
    public void TheDistanceIsWrittenTheWayTheReferenceWritesIt()
    {
        CultureInfo culture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

            Assert.AreEqual("1.2 km", SessionTracker.FormatDistance(1200),
                "Past a thousand metres it is kilometres, with one decimal.");
            Assert.AreEqual("340 m", SessionTracker.FormatDistance(340),
                "Below a thousand metres it is whole metres, never a decimal.");
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
        }
    }

    // ---- the saved file -----------------------------------------------------------------------

    [TestMethod]
    public void TheSessionIsWrittenAndReadBackFromTheFileTheReferenceUses()
    {
        using var folder = new Folder();
        var store = new LegacySessionTrackerStore(folder.Path);
        var state = new SessionState
        {
            ServerKey = ServerA,
            StartedUtc = Start.UtcDateTime,
            LastSeenUtc = Start.UtcDateTime,
            OnlineSeconds = 125.5,
            AfkSeconds = 12,
            DistanceMetres = 1234.5,
            Deaths = 3,
            TeamDeaths = 7,
        };
        state.Population.Add(new PopulationSample(Start.UtcDateTime, 42));

        store.Save(state);

        Assert.IsTrue(File.Exists(Path.Combine(folder.Path, "session_10_0_0_1_28082.json")),
            "The file is named after the server, the way the reference names it.");

        SessionState? read = store.Load(ServerA);
        Assert.IsNotNull(read);
        Assert.AreEqual(ServerA, read.ServerKey);
        Assert.AreEqual(125.5, read.OnlineSeconds, 0.001);
        Assert.AreEqual(12, read.AfkSeconds, 0.001);
        Assert.AreEqual(1234.5, read.DistanceMetres, 0.001);
        Assert.AreEqual(3, read.Deaths);
        Assert.AreEqual(7, read.TeamDeaths);
        Assert.AreEqual(Start.UtcDateTime, read.StartedUtc);
        Assert.AreEqual(1, read.Population.Count);
        Assert.AreEqual(42, read.Population[0].Players);
        Assert.AreEqual(Start.UtcDateTime, read.Population[0].At);
    }

    [TestMethod]
    public void TheFileUsesTheReferencePropertyNames()
    {
        using var folder = new Folder();
        var store = new LegacySessionTrackerStore(folder.Path);
        var state = new SessionState { ServerKey = ServerA, OnlineSeconds = 5 };
        state.Population.Add(new PopulationSample(Start.UtcDateTime, 42));
        store.Save(state);

        string json = File.ReadAllText(Path.Combine(folder.Path, "session_10_0_0_1_28082.json"));
        Assert.IsTrue(json.Contains("\"ServerKey\""), "The reference's own names, so both apps read the same file.");
        Assert.IsTrue(json.Contains("\"OnlineSeconds\""));
        Assert.IsTrue(json.Contains("\"DistanceMetres\""));
        Assert.IsTrue(json.Contains("\"Utc\""), "The population reading keeps the reference's own name.");
        Assert.IsFalse(json.Contains("\"At\""), "And not this app's own name for it.");
    }

    [TestMethod]
    public void AServerWithNoSessionHasNone()
    {
        using var folder = new Folder();
        var store = new LegacySessionTrackerStore(folder.Path);
        Assert.IsNull(store.Load(ServerA), "No file, no session.");
    }

    [TestMethod]
    public void ASessionFileThatCannotBeReadIsNoSession()
    {
        using var folder = new Folder();
        File.WriteAllText(Path.Combine(folder.Path, "session_10_0_0_1_28082.json"), "{ not json at all");
        var store = new LegacySessionTrackerStore(folder.Path);
        Assert.IsNull(store.Load(ServerA), "A file that cannot be read is a session that has not happened yet.");
    }
}
