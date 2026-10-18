using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Deaths;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeathTrackingTests
{
    private static readonly DeathLocationClassifier Open = new([], (_, _) => null, (_, _) => "G7");

    private static ServerTeamMember M(string id, bool alive, long? died = null, long? spawned = null, double? x = 10, double? y = 20) =>
        new(id, "Player " + id, true, alive) { DeathTime = died, SpawnTime = spawned, X = x, Y = y };

    [TestMethod]
    public void Tracker_FirstSnapshotIsHistory_ThenDeathTimeAdvancesAreDeaths()
    {
        var tracker = new DeathTracker();
        Assert.AreEqual(0, tracker.Observe([M("76561198000000001", false, died: 100)], Open, 1000).Count, "Deaths present on connect are history.");
        Assert.AreEqual(0, tracker.Observe([M("76561198000000001", false, died: 100)], Open, 1010).Count, "Staying dead is not a new death.");
        var deaths = tracker.Observe([M("76561198000000001", false, died: 200, spawned: 150)], Open, 1020);
        Assert.AreEqual(1, deaths.Count);
        Assert.AreEqual(new DeathRecord(76561198000000001, "Player 76561198000000001", 200, 150, 10, 20, "G7", "open", null), deaths[0]);
        Assert.AreEqual(0, tracker.Observe([M("0", false, died: 999), M("bad", false, died: 999)], Open, 1030).Count, "Invalid Steam IDs are skipped.");
    }

    [TestMethod]
    public void Tracker_WithoutDeathTime_UsesTheAliveToDeadEdgeAndTheObservedRespawn()
    {
        var tracker = new DeathTracker();
        tracker.Observe([M("1", true)], Open, 1000);
        tracker.Observe([M("1", false)], Open, 1100);
        tracker.Observe([M("1", true)], Open, 1200);
        var deaths = tracker.Observe([M("1", false)], Open, 1500);
        Assert.AreEqual(1, deaths.Count);
        Assert.AreEqual(1500, deaths[0].DeathTime, "Stamped with the current time.");
        Assert.AreEqual(1200, deaths[0].SpawnTime, "Survival is measured from the respawn the tracker saw.");
        Assert.AreEqual(0, tracker.Observe([M("1", false)], Open, 1600).Count);

        tracker.Reset();
        Assert.AreEqual(0, tracker.Observe([M("1", true)], Open, 1700).Count);
        Assert.AreEqual(1, tracker.Observe([M("1", false)], Open, 1800).Count, "A reset starts a new baseline.");
    }

    [TestMethod]
    public void Classifier_BaseWinsOverMonument_ThenNearestMonument_ElseOpen()
    {
        var team = new ServerTeamSnapshot([], [new TeamMapNote(2, 1000, 1000, "")], [new TeamMapNote(2, 3000, 3000, "Main"), new TeamMapNote(6, 500, 500, "Skull")]);
        var zones = new[] { new DeathZone(1050, 1000, 130, "Harbor"), new DeathZone(500, 500, 130, "Dome"), new DeathZone(560, 500, 130, "Closer") };
        var classifier = new DeathLocationClassifier(zones, (x, y) => DeathLocationClassifier.ResolveBase(team, x, y), (_, _) => "A1");
        Assert.AreEqual(("base", (string?)"Base"), classifier.Classify(1010, 1000), "A base wins over a monument; an unnamed note is \"Base\".");
        Assert.AreEqual(("base", (string?)"Main"), classifier.Classify(3089, 3000));
        Assert.AreEqual(("monument", (string?)"Closer"), classifier.Classify(540, 500), "Only Home notes are bases; the nearest zone wins.");
        Assert.AreEqual(("open", (string?)null), classifier.Classify(2000, 100));
        Assert.AreEqual(("open", (string?)null), classifier.Classify(null, 5));
        Assert.IsNull(classifier.Grid(1, null));
    }

    [TestMethod]
    public void MonumentZones_UseMapNames_SkipTrainTunnels_AndReferenceRadii()
    {
        var zones = DeathLocationClassifier.MonumentZones([
            new MapMonument("launchsite", 1, 1), new MapMonument("harbor_display_name", 2, 2), new MapMonument("train_tunnel_display_name", 3, 3),
            new MapMonument("supermarket", 4, 4)]);
        Assert.AreEqual(3, zones.Count, "Train tunnel entrances are skipped.");
        Assert.AreEqual(250, zones[0].Radius);
        Assert.AreEqual(160, zones[1].Radius);
        Assert.AreEqual(130, zones[2].Radius);
        Assert.AreEqual(MonumentFormatter.Beautify("harbor_display_name"), zones[1].Name);
        Assert.AreEqual(180, DeathLocationClassifier.MonumentRadiusFor("Water Treatment Plant"));
        Assert.AreEqual(120, DeathLocationClassifier.MonumentRadiusFor("large oil rig"));
        Assert.AreEqual(200, DeathLocationClassifier.MonumentRadiusFor("Airfield"));
    }

    [TestMethod]
    public void LogFile_WritesTheReferenceJsonLines_AndMergesTheLegacyFolder()
    {
        string root = Path.Combine(Path.GetTempPath(), "rpd-deaths-" + Guid.NewGuid().ToString("N"));
        string folder = Path.Combine(root, "RustPlusDesk", "deaths"), legacy = Path.Combine(root, "RustPlusDesktop", "deaths");
        try
        {
            Directory.CreateDirectory(legacy);
            File.WriteAllText(Path.Combine(legacy, "1.2.3.4-28082.jsonl"), "old-1\nshared\n");
            File.WriteAllText(Path.Combine(legacy, "5.6.7.8-1.jsonl"), "moved\n");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "1.2.3.4-28082.jsonl"), "shared\n");

            var log = new DeathLogFile(folder, legacy);
            log.Append("1.2.3.4-28082", new DeathRecord(76561198000000001, "Ann", 200, null, 1.5, 2, "A1", "monument", "Harbor"));
            var lines = File.ReadAllLines(log.PathFor("1.2.3.4-28082"));
            CollectionAssert.AreEqual(new[] { "shared", "old-1",
                "{\"steam_id\":\"76561198000000001\",\"name\":\"Ann\",\"died_at\":200,\"spawn_at\":null,\"x\":1.5,\"y\":2,\"grid\":\"A1\",\"location_type\":\"monument\",\"location_name\":\"Harbor\"}" },
                lines, "Legacy lines merge without duplicates, then the new death in the reference shape.");
            Assert.AreEqual("moved\n", File.ReadAllText(Path.Combine(folder, "5.6.7.8-1.jsonl")));
            Assert.IsFalse(Directory.Exists(Path.Combine(root, "RustPlusDesktop")), "The emptied legacy folders are removed.");
            StringAssert.EndsWith(log.PathFor("a/b"), "a_b.jsonl");
            StringAssert.EndsWith(log.PathFor(""), "server.jsonl");
        }
        finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [TestMethod]
    public void UnixSeconds_TreatsTheEpochAsNotSet()
    {
        Assert.IsNull(RustPlusDesk.Infrastructure.Connections.RustPlusTransport.UnixSeconds(DateTime.UnixEpoch));
        Assert.IsNull(RustPlusDesk.Infrastructure.Connections.RustPlusTransport.UnixSeconds(DateTime.MinValue));
        Assert.AreEqual(1_700_000_000, RustPlusDesk.Infrastructure.Connections.RustPlusTransport.UnixSeconds(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000).UtcDateTime));
        Assert.AreEqual(1_700_000_000, RustPlusDesk.Infrastructure.Connections.RustPlusTransport.UnixSeconds(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000).LocalDateTime));
    }

    [TestMethod]
    public async Task Recorder_FedByTheTeamLoop_LogsPerServerAndRebaselinesOnAServerChange()
    {
        var log = new MemoryLog();
        string key = "1.2.3.4-28082";
        var recorder = new DeathRecorder(log, () => key, () => [new MapMonument("harbor_display_name", 1000, 1000)], () => 4500, () => 5000);
        var team = new TeamViewModel();
        var snapshots = new List<ServerTeamSnapshot>();
        team.SnapshotRead += snapshot => { snapshots.Add(snapshot); recorder.Observe(snapshot); };
        var session = new ScriptedTeam([M("1", false, died: 100, x: 1010, y: 1000)]);
        using var cts = new CancellationTokenSource();
        var running = team.RunAsync(session, cts.Token);
        await CompanionListenerViewModelTests.Eventually(() => snapshots.Count == 1);
        cts.Cancel();
        await running;
        Assert.AreEqual(0, snapshots[0].Notes.Count + snapshots[0].LeaderNotes.Count, "The default snapshot has no notes.");
        Assert.AreEqual(0, log.Entries.Count, "The first snapshot is the baseline.");

        recorder.Observe(new([M("1", false, died: 200, x: 1010, y: 1000)], [], []));
        Assert.AreEqual(1, log.Entries.Count);
        Assert.AreEqual(key, log.Entries[0].Key);
        Assert.AreEqual("monument", log.Entries[0].Death.LocationType, "The harbor zone classifies it.");
        Assert.AreEqual(MapGrid.Label(4500, 1010, 1000), log.Entries[0].Death.Grid);

        key = "9.9.9.9-1";
        Assert.AreEqual(0, recorder.Observe(new([M("1", false, died: 300)], [], [])).Count, "A new server starts from a fresh baseline.");
        Assert.AreEqual(1, recorder.Observe(new([M("1", false, died: 400)], [], [])).Count);
        Assert.AreEqual("9.9.9.9-1", log.Entries[^1].Key);
        Assert.AreEqual("1.2.3.4-28082", DeathRecorder.KeyFor("1.2.3.4", 28082));
    }

    [TestMethod]
    public async Task Shell_RecordsDeathsForTheSelectedServer()
    {
        var log = new MemoryLog();
        var shell = new RustPlusDesk.Avalonia.Features.Servers.ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory(), deathLog: log);
        await shell.RefreshCommand.ExecuteAsync(null);
        Assert.IsNotNull(shell.Deaths);
        shell.Deaths.Observe(new([M("1", true)], [], []));
        shell.Deaths.Observe(new([M("1", false, died: 50)], [], []));
        Assert.AreEqual("1.2.3.4-28082", log.Entries.Single().Key);
        Assert.AreEqual("off-grid", log.Entries[0].Death.Grid, "Without a loaded map the reference grid label is off-grid.");
        Assert.IsNull(new RustPlusDesk.Avalonia.Features.Servers.ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory()).Deaths);
    }

    private sealed class OneServer : IServerCatalog
    {
        public Task<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>>([new("Rusty", "1.2.3.4", 28082) { ProfileKey = "key" }]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    private sealed class MemoryLog : IDeathLog
    {
        public List<(string Key, DeathRecord Death)> Entries { get; } = [];
        public void Append(string serverKey, DeathRecord death) => Entries.Add((serverKey, death));
        public IReadOnlyList<string> ReadLines(string serverKey) => [];
        public void Clear(string serverKey) { }
        public IReadOnlyList<string> ReadKillerLines(string serverKey) => [];
        public void WriteKillerLines(string serverKey, IEnumerable<string> lines) { }
        public void ClearKillers(string serverKey) { }
    }

    /// <summary>A team session without map notes: the interface's default snapshot wraps ReadTeamAsync.</summary>
    private sealed class ScriptedTeam(params IReadOnlyList<ServerTeamMember>[] reads) : IServerTeamSession
    {
        public int Reads;
        public RustPlusDesk.Core.Devices.IDeviceClient Devices => throw new NotSupportedException();
        public ServerStatus InitialStatus => new("Test", 1, 1, 0, 12);
        public Task Disconnected => Task.Delay(Timeout.Infinite);
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct) => Task.FromResult(reads[Math.Min(Interlocked.Increment(ref Reads) - 1, reads.Length - 1)]);
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
