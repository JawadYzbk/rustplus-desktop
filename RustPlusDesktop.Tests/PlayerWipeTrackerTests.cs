using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Services.PlayerWipeTracker;

namespace RustPlusDesktop.Tests;

[TestClass]
public sealed class PlayerWipeTrackerTests
{
    private static PlayerObservation Observation(
        DateTime timestamp,
        bool online = true,
        bool dead = false,
        bool afk = false,
        double? x = 0,
        double? y = 0,
        TrackerLocationType location = TrackerLocationType.Open,
        string? locationName = null,
        string session = "s1")
        => new(76561198000000001, "Player", timestamp.ToUniversalTime(), session, true, true, online, dead, afk, x, y, location, locationName, "A1", null, null);

    [TestMethod]
    public void FirstSnapshot_EstablishesBaselineWithoutElapsedTime()
    {
        var engine = new PlayerWipeTrackerEngine();

        engine.Observe(Observation(DateTime.UtcNow));

        Assert.AreEqual(TimeSpan.Zero, engine.Summarize().Coverage);
    }

    [TestMethod]
    public void StatePriority_UnknownOfflineDeadAfkMovingAndStationary()
    {
        var now = DateTime.UtcNow;
        Assert.AreEqual(PlayerActivityState.Unknown, PlayerWipeTrackerEngine.Classify(Observation(now) with { SnapshotValid = false }));
        Assert.AreEqual(PlayerActivityState.Offline, PlayerWipeTrackerEngine.Classify(Observation(now, online: false)));
        Assert.AreEqual(PlayerActivityState.Dead, PlayerWipeTrackerEngine.Classify(Observation(now, dead: true)));
        Assert.AreEqual(PlayerActivityState.Afk, PlayerWipeTrackerEngine.Classify(Observation(now, afk: true)));
        Assert.AreEqual(PlayerActivityState.Stationary, PlayerWipeTrackerEngine.Classify(Observation(now)));
        Assert.AreEqual(PlayerActivityState.Moving, PlayerWipeTrackerEngine.Classify(Observation(now, x: 20), 20));
    }

    [TestMethod]
    public void ReconnectGap_IsUnknownAndDoesNotAddDistance()
    {
        var engine = new PlayerWipeTrackerEngine();
        var start = DateTime.UtcNow;
        engine.Observe(Observation(start, x: 0));
        engine.Observe(Observation(start.AddSeconds(5), x: 5));
        engine.Observe(Observation(start.AddMinutes(2), x: 500, session: "s2"));

        var summary = engine.Summarize();
        Assert.IsTrue(summary.Unknown >= TimeSpan.FromMinutes(1));
        Assert.IsTrue(summary.EstimatedDistance < 100);
    }

    [TestMethod]
    public async Task JsonLinesStore_SkipsCorruptLinesAndDeduplicates()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"tracker-{Guid.NewGuid():N}");
        await using var store = new PlayerWipeTrackerStore(directory);
        var observation = Observation(DateTime.UtcNow);
        var item = new TrackerPersistedObservation(1, "observation", observation);
        Assert.IsTrue(store.Append("server", "wipe", observation.SteamId, item));
        Assert.IsTrue(store.Append("server", "wipe", observation.SteamId, item));
        await store.FlushAsync();

        var path = Directory.EnumerateFiles(directory, "*.jsonl", SearchOption.AllDirectories).Single();
        await File.AppendAllTextAsync(path, "not json\n");

        var loaded = store.Load("server", "wipe", observation.SteamId);
        Assert.AreEqual(1, loaded.Count);
        store.DeleteAll();
        Assert.AreEqual(0, store.StorageBytes);
    }
}
