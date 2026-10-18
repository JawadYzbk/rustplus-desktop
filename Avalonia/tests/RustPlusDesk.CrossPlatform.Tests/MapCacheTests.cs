using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Infrastructure.Map;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class MapCacheTests
{
    private static readonly DateTimeOffset Wipe = new(2026, 10, 1, 12, 0, 0, TimeSpan.Zero);

    [TestMethod]
    public void WipeRules_FollowTheReference()
    {
        Assert.IsFalse(MapWipe.IsNewerWipe(null, Wipe), "Without a server wipe time nothing is newer.");
        Assert.IsTrue(MapWipe.IsNewerWipe(Wipe, null));
        Assert.IsFalse(MapWipe.IsNewerWipe(Wipe.AddSeconds(59), Wipe), "Within a minute is the same wipe.");
        Assert.IsTrue(MapWipe.IsNewerWipe(Wipe.AddMinutes(2), Wipe));
        Assert.IsFalse(MapWipe.IsNewerWipe(Wipe, Wipe.AddDays(1)));

        var harbors = MapWipe.Harbors([new("harbor_2_display_name", 10, 10), new("large_oil_rig", 1, 1), new("Harbor_1_display_name", 20, 20)]);
        CollectionAssert.AreEqual(new[] { "harbor_2_display_name", "Harbor_1_display_name" }.OrderBy(name => name, StringComparer.CurrentCulture).ToArray(),
            harbors.Select(h => h.Name).ToArray());
        Assert.IsFalse(MapWipe.HarborsChanged(harbors, [.. harbors.Reverse().Select(h => h with { X = h.X + 50, Y = h.Y - 50 })]), "50 m is tolerated.");
        Assert.IsTrue(MapWipe.HarborsChanged(harbors, [.. harbors.Select(h => h with { X = h.X + 50.5 })]));
        Assert.IsTrue(MapWipe.HarborsChanged(harbors, [harbors[0]]));
        Assert.IsTrue(MapWipe.HarborsChanged(harbors, [harbors[0], harbors[1] with { Name = "other harbor" }]));
        Assert.IsFalse(MapWipe.IsWipeDetected([], harbors), "No harbors on the map cannot be judged.");
        Assert.IsFalse(MapWipe.IsWipeDetected(harbors, []), "The first connection is not a wipe.");
        Assert.IsTrue(MapWipe.IsWipeDetected(harbors, [harbors[0]]));
    }

    [TestMethod]
    public void Store_UsesTheReferenceNamesAndJson_AndDeletesTheCustomImageToo()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-map-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            var store = new MapCacheStore(folder);
            Assert.AreEqual("10.0.0.1_28082", MapCacheStore.Key("10.0.0.1", 28082));
            Assert.AreEqual("a_b_c_d_1", MapCacheStore.Key("a:b/c?d", 1));
            Assert.IsNull(store.Load("10.0.0.1", 28082));

            store.Save("10.0.0.1", 28082, new CachedMap([1, 2, 3], 2, 3, 4000, Wipe, [new("harbor_display_name", 1.5, 2)]));
            Assert.AreEqual("{\"WorldSize\":4000,\"PixelWidth\":2,\"PixelHeight\":3,\"WipeTime\":\"2026-10-01T12:00:00Z\",\"Monuments\":[{\"X\":1.5,\"Y\":2,\"Name\":\"harbor_display_name\"}]}",
                File.ReadAllText(Path.Combine(folder, "10.0.0.1_28082.json")));
            var loaded = store.Load("10.0.0.1", 28082)!;
            CollectionAssert.AreEqual(new byte[] { 1, 2, 3 }, loaded.Image);
            Assert.AreEqual((2, 3, 4000, Wipe), (loaded.PixelWidth, loaded.PixelHeight, loaded.WorldSize, loaded.WipeTime));
            Assert.AreEqual(new MapMonument("harbor_display_name", 1.5, 2), loaded.Monuments.Single());
            var server = loaded.ToServerMap();
            Assert.AreEqual((2, 3, 0, 4000u), (server.Width, server.Height, server.OceanMargin, server.WorldSize));

            File.WriteAllText(Path.Combine(folder, "10.0.0.1_28082.json"), "{\"WorldSize\":3000,\"PixelWidth\":1,\"PixelHeight\":1,\"WipeTime\":null,\"Monuments\":[]}");
            Assert.IsNull(store.Load("10.0.0.1", 28082)!.WipeTime);
            File.WriteAllText(Path.Combine(folder, "10.0.0.1_28082.json"), "{ broken");
            Assert.IsNull(store.Load("10.0.0.1", 28082), "Unreadable metadata is no cache.");

            File.WriteAllBytes(Path.Combine(folder, "10.0.0.1_28082_custom.png"), [9]);
            File.WriteAllBytes(Path.Combine(folder, "10.0.0.2_28082.png"), [9]);
            store.Delete("10.0.0.1", 28082);
            CollectionAssert.AreEqual(new[] { "10.0.0.2_28082.png" }, Directory.GetFiles(folder).Select(Path.GetFileName).ToArray(), "Other servers keep their cache.");
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public async Task Loader_UsesTheCacheUntilAWipe_SavesPng_AndRemembersHarbors() => await UiTestSession.Current.Dispatch(async () =>
    {
        var cache = new MemoryCache();
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(MapWipe.ServerCargoTriggersKey, new Dictionary<string, JsonElement> { ["10.0.0.1"] = JsonSerializer.SerializeToElement(new { }), ["10.0.0.9"] = JsonSerializer.SerializeToElement(new { }) });
        var log = new List<string>();
        int wipes = 0;
        var loader = new MapCacheLoader(cache, settings, () => ("10.0.0.1", 28082)) { Log = log.Add };
        loader.WipeDetected += () => wipes++;
        var synthetic = MapTests.SyntheticMap();
        var map = synthetic with { Monuments = [.. synthetic.Monuments, new("harbor_display_name", 1000, 1000)] };

        var session = new Session(map, Wipe);
        var first = await loader.LoadAsync(session, default);
        Assert.AreEqual(1, session.Calls);
        Assert.AreSame(map, first);
        var saved = cache.Entries["10.0.0.1:28082"];
        CollectionAssert.AreEqual(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, saved.Image.Take(4).ToArray(), "The reference stores PNG.");
        Assert.AreEqual((1024, 1024, 4500, Wipe), (saved.PixelWidth, saved.PixelHeight, saved.WorldSize, saved.WipeTime));
        var harbors = settings.Get<Dictionary<string, List<HarborInfo>>>(MapWipe.ServerHarborsKey, [])["10.0.0.1"];
        Assert.AreEqual(new HarborInfo("harbor_display_name", 1000, 1000), harbors.Single());
        CollectionAssert.AreEqual(new[] { "10.0.0.9" }, settings.Get<Dictionary<string, JsonElement>>(MapWipe.ServerCargoTriggersKey, []).Keys.ToArray(),
            "New harbors clear the host's cargo triggers.");

        var cached = await loader.LoadAsync(new Session(map, Wipe.AddSeconds(30)), default);
        Assert.AreEqual(4, cached.Monuments.Count);
        Assert.AreEqual(0, wipes);
        Assert.IsTrue(log.Any(line => line.StartsWith("[map-cache] Loaded map from disk cache", StringComparison.Ordinal)));

        var newWipe = new Session(map, Wipe.AddDays(7));
        await loader.LoadAsync(newWipe, default);
        Assert.AreEqual((1, 1, 1), (newWipe.Calls, wipes, cache.Deletes), "A newer server wipe invalidates the cache.");
        Assert.AreEqual(Wipe.AddDays(7), cache.Entries["10.0.0.1:28082"].WipeTime);

        var moved = map with { Monuments = [new("harbor_display_name", 3000, 3000)] };
        var noWipeTime = new Session(moved, DateTimeOffset.UnixEpoch);
        await loader.LoadAsync(noWipeTime, default);
        Assert.AreEqual((0, 1), (noWipeTime.Calls, wipes), "Same harbors without a server wipe time: the cache stays.");
        cache.Entries["10.0.0.1:28082"] = cache.Entries["10.0.0.1:28082"] with { Monuments = [new("harbor_display_name", 3000, 3000)] };
        await loader.LoadAsync(noWipeTime, default);
        Assert.AreEqual((1, 2), (noWipeTime.Calls, wipes), "Moved harbors without a server wipe time are a wipe.");
        Assert.IsTrue(log.Contains("[map-cache] Wipe detected from monument layout; invalidating map cache."));
        Assert.IsNull(cache.Entries["10.0.0.1:28082"].WipeTime);

        var noServer = new MapCacheLoader(cache, settings, () => null);
        var direct = new Session(map, Wipe);
        await noServer.LoadAsync(direct, default);
        Assert.AreEqual(1, direct.Calls);
        return true;
    }, default);

    private sealed class MemoryCache : IMapCache
    {
        public Dictionary<string, CachedMap> Entries { get; } = [];
        public int Deletes { get; private set; }
        public CachedMap? Load(string host, int port) => Entries.GetValueOrDefault($"{host}:{port}");
        public void Save(string host, int port, CachedMap map) => Entries[$"{host}:{port}"] = map;
        public void Delete(string host, int port) { Deletes++; Entries.Remove($"{host}:{port}"); }
    }

    private sealed class Session(ServerMap map, DateTimeOffset? wipe) : IServerMapSession
    {
        public int Calls { get; private set; }
        public IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Synthetic", 0, 0, 0, WipeTime: wipe);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public Task<ServerMap> ReadMapAsync(CancellationToken ct) { Calls++; return Task.FromResult(map); }
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
