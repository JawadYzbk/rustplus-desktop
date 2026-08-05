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
    public void MapProjection_AlignsWorldCornersWithPaddedUniformImage()
    {
        var projection = new TrackerMapProjection(
            ViewWidth: 800,
            ViewHeight: 500,
            ImageWidth: 1000,
            ImageHeight: 1000,
            WorldRectX: 100,
            WorldRectY: 100,
            WorldRectWidth: 800,
            WorldRectHeight: 800,
            WorldSize: 4000);

        var northWest = projection.Project(0, 4000);
        var southEast = projection.Project(4000, 0);

        Assert.AreEqual(200, northWest.X, 0.001);
        Assert.AreEqual(50, northWest.Y, 0.001);
        Assert.AreEqual(600, southEast.X, 0.001);
        Assert.AreEqual(450, southEast.Y, 0.001);
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

    [TestMethod]
    public async Task Store_KeepsMapInsideItsWipeDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"tracker-map-{Guid.NewGuid():N}");
        await using var store = new PlayerWipeTrackerStore(directory);
        var expected = new TrackerWipeMap(new byte[] { 1, 2, 3 }, 4500, 10, 20, 900, 900);

        store.SaveWipeMap("server", "wipe-a", expected);

        var actual = store.LoadWipeMap("server", "wipe-a");
        Assert.IsNotNull(actual);
        CollectionAssert.AreEqual(expected.PngBytes, actual.PngBytes);
        Assert.AreEqual(expected.WorldSize, actual.WorldSize);
        Assert.AreEqual(expected.WorldRectWidth, actual.WorldRectWidth);
        Assert.IsNull(store.LoadWipeMap("server", "wipe-b"));
        store.DeleteAll();
    }
}
