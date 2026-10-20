using System.Text;
using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Tracking;
using RustPlusDesk.Infrastructure.Tracking;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference PlayerWipeTrackerRecordFile and PlayerWipeTrackerStore: the observations of the tracked players.
///
/// The record file is compact - the strings that repeat are written once and referred to by number afterwards - and
/// every record is self-delimiting, so a record cut short by a crash ends the readable part of the file and the next
/// append cuts it away.
/// </summary>
[TestClass]
public sealed class PlayerWipeTrackerTests
{
    private static readonly DateTime Now = new(2026, 10, 20, 12, 0, 0, DateTimeKind.Utc);
    private const ulong SteamId = 76561198000000001;

    private static PlayerObservation Observation(string name = "Alice", DateTime? at = null, double? x = 100.5, double? y = -200.25) =>
        new(SteamId, name, at ?? Now, "session-1", IsConnected: true, SnapshotValid: true, Online: true, Dead: false, Afk: false,
            x, y, TrackerLocationType.Monument, "Launch Site", "C12", 1_700_000_000, null);

    // ---- the record file -----------------------------------------------------------------------

    [TestMethod]
    public void TheObservationIsReadBackAsItWasWritten()
    {
        var file = TrackerRecordFile.CreateNew(Now);
        byte[] bytes = file.Encode(Observation());
        var observations = new List<PlayerObservation>();
        var read = TrackerRecordFile.Parse(TrackerRecordFile.Header(file.BaseTicks).Concat(bytes).ToArray(), SteamId, observations);

        Assert.IsNotNull(read, "The file is this format.");
        Assert.AreEqual(1, observations.Count);
        var observation = observations[0];
        Assert.AreEqual(SteamId, observation.SteamId);
        Assert.AreEqual("Alice", observation.Name);
        Assert.AreEqual(Now, observation.TimestampUtc);
        Assert.AreEqual("session-1", observation.SessionId);
        Assert.IsTrue(observation.IsConnected);
        Assert.IsTrue(observation.SnapshotValid);
        Assert.IsTrue(observation.Online);
        Assert.IsFalse(observation.Dead);
        Assert.IsFalse(observation.Afk);
        Assert.AreEqual(100.5, observation.X!.Value, 0.001);
        Assert.AreEqual(-200.25, observation.Y!.Value, 0.001);
        Assert.AreEqual(TrackerLocationType.Monument, observation.LocationType);
        Assert.AreEqual("Launch Site", observation.LocationName);
        Assert.AreEqual("C12", observation.Grid);
        Assert.AreEqual(1_700_000_000, observation.SpawnTime);
        Assert.IsNull(observation.DeathTime);
    }

    [TestMethod]
    public void TheStringsThatRepeatAreWrittenOnce()
    {
        var file = TrackerRecordFile.CreateNew(Now);
        byte[] first = file.Encode(Observation());
        byte[] second = file.Encode(Observation(at: Now.AddSeconds(30)));

        // The name, the session, the monument and the grid are written once and referred to afterwards.
        Assert.IsTrue(second.Length < first.Length / 2, "The second record only refers to the strings it already knows.");
        var observations = new List<PlayerObservation>();
        TrackerRecordFile.Parse(TrackerRecordFile.Header(file.BaseTicks).Concat(first).Concat(second).ToArray(), SteamId, observations);
        Assert.AreEqual(2, observations.Count);
        Assert.AreEqual("Alice", observations[1].Name);
        Assert.AreEqual(Now.AddSeconds(30), observations[1].TimestampUtc);
    }

    [TestMethod]
    public void ARecordThatWasCutShortEndsTheReadablePart()
    {
        var file = TrackerRecordFile.CreateNew(Now);
        byte[] bytes = file.Encode(Observation());
        byte[] torn = bytes[..(bytes.Length - 4)];
        var observations = new List<PlayerObservation>();
        var read = TrackerRecordFile.Parse(TrackerRecordFile.Header(file.BaseTicks).Concat(torn).ToArray(), SteamId, observations);

        Assert.IsNotNull(read, "The header is still there.");
        Assert.AreEqual(0, observations.Count, "The torn record is not read.");
        Assert.IsTrue(read!.ValidLength < TrackerRecordFile.HeaderLength + torn.Length,
            "The file ends where the last complete record ended.");
    }

    [TestMethod]
    public void BytesThatAreNotThisFormatAreNotRead()
    {
        Assert.IsNull(TrackerRecordFile.Parse(Encoding.UTF8.GetBytes("not a tracker file"), SteamId, null));
        Assert.IsFalse(TrackerRecordFile.HasHeader(Encoding.UTF8.GetBytes("PWT")));
    }

    [TestMethod]
    public void TheLocationTypeIsReadAsTheNumberItIsStoredAs()
    {
        var file = TrackerRecordFile.CreateNew(Now);
        var observation = Observation() with { LocationType = TrackerLocationType.DeepSea, LocationName = null, Grid = null };
        byte[] bytes = file.Encode(observation);
        var observations = new List<PlayerObservation>();
        TrackerRecordFile.Parse(TrackerRecordFile.Header(file.BaseTicks).Concat(bytes).ToArray(), SteamId, observations);
        Assert.AreEqual(TrackerLocationType.DeepSea, observations[0].LocationType);
        Assert.IsNull(observations[0].LocationName);
        Assert.AreEqual("deep_sea", TrackerLocationType.DeepSea.Wire(), "The wire's spelling is its own.");
    }

    // ---- the store -----------------------------------------------------------------------------

    [TestMethod]
    public async Task TheObservationsAreWrittenAndReadBack()
    {
        using var folder = new TempFolder();
        using var store = new PlayerWipeTrackerStore(folder.Path);
        await store.AppendAsync("10.0.0.1:28082", "wipe-1", SteamId, Observation());
        await store.AppendAsync("10.0.0.1:28082", "wipe-1", SteamId, Observation(at: Now.AddMinutes(1)));

        var read = store.Load("10.0.0.1:28082", "wipe-1", SteamId);
        Assert.AreEqual(2, read.Count);
        Assert.AreEqual("Alice", read[0].Name);
        Assert.AreEqual(Now.AddMinutes(1), read[1].TimestampUtc);
        Assert.IsTrue(File.Exists(Path.Combine(folder.Path, "10.0.0.1_28082", "wipe-1", $"{SteamId}.pwt")),
            "The file is where the reference keeps it.");
    }

    [TestMethod]
    public async Task TheHistoryTheAppWroteAsJsonLinesIsConverted()
    {
        using var folder = new TempFolder();
        var directory = Path.Combine(folder.Path, "10.0.0.1_28082", "wipe-1");
        Directory.CreateDirectory(directory);
        var legacy = Path.Combine(directory, $"{SteamId}.jsonl");
        var item = new TrackerPersistedObservation(1, "observation", Observation());
        await File.WriteAllTextAsync(legacy, JsonSerializer.Serialize(item) + Environment.NewLine
            + "{ this line is broken" + Environment.NewLine);

        using var store = new PlayerWipeTrackerStore(folder.Path);
        var read = store.Load("10.0.0.1:28082", "wipe-1", SteamId);
        Assert.AreEqual(1, read.Count, "The broken line does not hide the history.");
        Assert.AreEqual("Alice", read[0].Name);
        Assert.IsFalse(File.Exists(legacy), "The old file is gone once it is converted.");
        Assert.IsTrue(File.Exists(Path.Combine(directory, $"{SteamId}.pwt")), "The history is in the record file now.");
    }

    [TestMethod]
    public async Task TheWipeKeepsItsMetadataAndItsMap()
    {
        using var folder = new TempFolder();
        using var store = new PlayerWipeTrackerStore(folder.Path);
        store.SaveWipeMetadata("10.0.0.1:28082", "wipe-1",
            new StoredWipeMetadata("10.0.0.1:28082", "Test Server", "wipe-1", Now, Now));
        await store.AppendAsync("10.0.0.1:28082", "wipe-1", SteamId, Observation());
        store.SaveWipeMap("10.0.0.1:28082", "wipe-1",
            new TrackerWipeMap([1, 2, 3], 4000, -1000, -1000, 6000, 6000));

        var metadata = store.LoadWipeMetadata("10.0.0.1:28082", "wipe-1");
        Assert.IsNotNull(metadata);
        Assert.AreEqual("Test Server", metadata!.ServerName);
        Assert.IsTrue(store.HasWipeMap("10.0.0.1:28082", "wipe-1"), "The wipe has a map.");
        var map = store.LoadWipeMap("10.0.0.1:28082", "wipe-1");
        Assert.IsNotNull(map);
        CollectionAssert.AreEqual(new byte[] { 1, 2, 3 }, map!.PngBytes, "The map's image is kept.");

        var wipes = store.GetStoredWipes();
        Assert.AreEqual(1, wipes.Count, "The store holds one wipe.");
        Assert.AreEqual("Test Server", wipes[0].ServerName);
        Assert.AreEqual(1, wipes[0].PlayerCount, "The wipe holds one player.");
        Assert.IsTrue(wipes[0].HasMap);
    }

    [TestMethod]
    public void AWipeIsGoneWhenItIsDeleted()
    {
        using var folder = new TempFolder();
        using var store = new PlayerWipeTrackerStore(folder.Path);
        store.SaveWipeMetadata("10.0.0.1:28082", "wipe-1",
            new StoredWipeMetadata("10.0.0.1:28082", "Test Server", "wipe-1", Now, Now));
        Assert.IsNotNull(store.LoadWipeMetadata("10.0.0.1:28082", "wipe-1"));

        store.DeleteWipe("10.0.0.1:28082", "wipe-1");
        Assert.IsNull(store.LoadWipeMetadata("10.0.0.1:28082", "wipe-1"), "The wipe is gone.");
        Assert.AreEqual(0, store.GetStoredWipes().Count);
    }

    /// <summary>A folder of its own, removed when the test is over.</summary>
    private sealed class TempFolder : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "rustplus-wipes-" + Guid.NewGuid().ToString("N"));

        public TempFolder() => Directory.CreateDirectory(Path);

        public void Dispose()
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
        }
    }

    /// <summary>Reference PlayerWipeTrackerEnabled: the tracker records nothing until the user turns it on.</summary>
    [TestMethod]
    public void TheTrackerIsOptInAndTheCloudBackupIsASeparateChoice()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var model = new RustPlusDesk.Avalonia.Features.Settings.SettingsViewModel(settings);
        Assert.IsFalse(model.PlayerWipeTrackerEnabled, "The tracker is opt-in.");
        Assert.IsFalse(model.PlayerWipeTrackerCloudBackup, "The cloud backup is a separate choice.");

        model.PlayerWipeTrackerEnabled = true;
        model.PlayerWipeTrackerCloudBackup = true;
        Assert.IsTrue(settings.Get(RustPlusDesk.Core.Settings.SettingKeys.PlayerWipeTrackerEnabled, false));
        Assert.IsTrue(settings.Get(RustPlusDesk.Core.Settings.SettingKeys.PlayerWipeTrackerCloudBackupEnabled, false));
    }
}
