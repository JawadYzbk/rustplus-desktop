using System.Text.Json.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ProfileWipeTimeTests
{
    private static readonly DateTimeOffset Wipe = new(2026, 10, 1, 12, 0, 0, TimeSpan.Zero);

    [TestMethod]
    public async Task Store_WritesTheReferenceUtcShape_OnlyWhenItChanges()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyProfileWipeTimeStore();
        Assert.IsNull(await store.GetAsync(fixture.Path, server.ProfileKey, default));
        var before = await fixture.JsonAsync();
        Assert.IsTrue(await store.SetAsync(fixture.Path, server.ProfileKey, Wipe.ToOffset(TimeSpan.FromHours(2)), default));
        var after = await fixture.JsonAsync();
        Assert.AreEqual("\"2026-10-01T12:00:00Z\"", after[0]!["WipeTime"]!.ToJsonString());
        after[0]!.AsObject().Remove("WipeTime");
        Assert.IsTrue(JsonNode.DeepEquals(before, after));
        Assert.AreEqual(Wipe, await store.GetAsync(fixture.Path, server.ProfileKey, default));
        string revision = (await fixture.ServerAsync()).SourceRevision;
        Assert.IsFalse(await store.SetAsync(fixture.Path, server.ProfileKey, Wipe, default), "The same wipe is not written again.");
        Assert.AreEqual(revision, (await fixture.ServerAsync()).SourceRevision);
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => store.SetAsync(fixture.Path, "missing", Wipe, default));
    }

    [TestMethod]
    public async Task Loader_SavesTheServersWipe_AndFallsBackToTheProfilesWipe() => await UiTestSession.Current.Dispatch(async () =>
    {
        var cache = new MapCacheTests.MemoryCache();
        DateTimeOffset? profileWipe = null;
        var saved = new List<DateTimeOffset>();
        var loader = new MapCacheLoader(cache, null, () => ("10.0.0.1", 28082))
        {
            ProfileWipeTime = _ => Task.FromResult(profileWipe),
            SaveProfileWipeTime = (wipe, _) => { saved.Add(wipe); profileWipe = wipe; return Task.CompletedTask; }
        };
        var map = MapTests.SyntheticMap();
        await loader.LoadAsync(new MapCacheTests.Session(map, Wipe), default);
        CollectionAssert.AreEqual(new[] { Wipe }, saved.ToArray());

        profileWipe = Wipe.AddDays(7);
        var noWipe = new MapCacheTests.Session(map, null);
        await loader.LoadAsync(noWipe, default);
        Assert.AreEqual(1, noWipe.Calls, "Without a server wipe time the profile's newer wipe invalidates the cache.");
        Assert.AreEqual(Wipe.AddDays(7), cache.Entries["10.0.0.1:28082"].WipeTime);
        Assert.AreEqual(1, saved.Count, "Nothing to save when the server reports no wipe.");
        return true;
    }, default);
}
