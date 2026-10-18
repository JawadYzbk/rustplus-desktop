using System.Globalization;
using System.Text.Json.Nodes;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Core.Settings;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeathPinTests
{
    private static readonly DateTime T0 = new(2026, 10, 18, 14, 5, 0, DateTimeKind.Local);

    private static DeathMarker D(ulong steamId, int minutes, double x = 1000, string name = "Player", string custom = "") =>
        new(Guid.NewGuid(), steamId, name, custom, x, 2000, T0.AddMinutes(minutes));

    [TestMethod]
    public void Rules_TrimPerPlayer_HideDeepSea_AndNumberOldestFirst()
    {
        var mine = new[] { D(1, 0), D(1, 1), D(1, 2) };
        var bob = new[] { D(2, 0), D(2, 5) };
        var kept = DeathMarkerRules.Trim([.. mine, .. bob], self: 1, maxSelf: 2, maxTeam: 1);
        CollectionAssert.AreEquivalent(new[] { mine[1], mine[2], bob[1] }, kept.ToArray(), "The oldest of each player go first.");
        var unknownSelf = DeathMarkerRules.Trim([.. mine, .. bob], self: null, maxSelf: 50, maxTeam: 1);
        CollectionAssert.AreEquivalent(new[] { mine[2], bob[1] }, unknownSelf.ToArray(), "Without an own Steam ID everybody has the team cap.");
        var added = DeathMarkerRules.Add([.. mine], D(1, 9), 1, 3, 3);
        Assert.AreEqual(3, added.Count);
        Assert.IsFalse(added.Contains(mine[0]));
        Assert.AreEqual(1, DeathMarkerRules.ClampCap(0));
        Assert.AreEqual(50, DeathMarkerRules.ClampCap(99));

        var deepSea = D(2, 9, x: -5000);
        var pins = DeathMarkerRules.Pins([bob[1], bob[0], deepSea, D(3, 1, name: "Cat", custom: "Base raid")], CultureInfo.InvariantCulture);
        CollectionAssert.AreEqual(new[] { "1. Player (14:05)", "2. Player (14:10)", "1. Base raid (14:06)" }, pins.Select(pin => pin.Label).ToArray());
        Assert.IsFalse(pins.Any(pin => pin.Marker == deepSea), "Deep Sea markers stay off the island map.");
    }

    [TestMethod]
    public async Task Store_WritesTheReferenceShape_KeepsEverythingElse_AndRejectsAMissingProfile()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDeathMarkerStore();
        Assert.AreEqual(0, (await store.LoadAsync(fixture.Path, server.ProfileKey, default)).Count);
        var before = await fixture.JsonAsync();
        var marker = new DeathMarker(Guid.Parse("0f8fad5b-d9cb-469f-a165-70867728950e"), 76561198000000001, "Ann", "", 1200.5, 800, new DateTime(2026, 10, 18, 12, 0, 0, DateTimeKind.Utc));
        var saved = await store.UpdateAsync(fixture.Path, server.ProfileKey, list => [.. list, marker], default);
        Assert.AreEqual(marker, saved.Single());
        var after = await fixture.JsonAsync();
        Assert.AreEqual("[{\"Id\":\"0f8fad5b-d9cb-469f-a165-70867728950e\",\"SteamId\":76561198000000001,\"OriginalName\":\"Ann\",\"CustomName\":\"\",\"X\":1200.5,\"Y\":800,\"TimeOfDeath\":\"2026-10-18T12:00:00Z\"}]",
            after[0]!["DeathMarkers"]!.ToJsonString());
        after[0]!.AsObject().Remove("DeathMarkers");
        Assert.IsTrue(JsonNode.DeepEquals(before, after), "Every other field, unknown ones included, is kept.");
        Assert.AreEqual(marker, (await store.LoadAsync(fixture.Path, server.ProfileKey, default)).Single());
        await store.UpdateAsync(fixture.Path, server.ProfileKey, list => list.Select(m => m with { CustomName = "Bag" }).ToArray(), default);
        Assert.AreEqual("Bag", (await store.LoadAsync(fixture.Path, server.ProfileKey, default)).Single().DisplayName);
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => store.UpdateAsync(fixture.Path, "missing", list => list, default));
    }

    [TestMethod]
    public async Task ViewModel_PinsTeamDeaths_AppliesCaps_RenamesDeletesAndWipes() => await UiTestSession.Current.Dispatch(async () =>
    {
        var store = new MemoryMarkers();
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.MaxSelfDeathMarkers, 1);
        var map = new MapViewModel();
        var team = new TeamViewModel();
        var log = new List<string>();
        var model = new DeathPinsViewModel(store, settings, team, map, () => ("profiles.json", "key"), () => "1") { Log = log.Add };
        int minute = 0;
        model.Now = () => T0.AddMinutes(minute++);
        Assert.AreEqual((false, 1, 3), (model.Show, model.MaxSelf, model.MaxTeam));

        ServerTeamMember M(string id, string name, bool alive, double? x = null, double? y = null) => new(id, name, true, alive) { X = x, Y = y };
        await model.ObserveAsync([M("1", "Me", true), M("2", "Bob", true)]);
        await model.ObserveAsync([M("1", "Me", false, 100, 200), M("2", "Bob", false)]);
        Assert.AreEqual(1, model.Markers.Count, "Bob has no position and no player marker: no pin.");
        Assert.AreEqual((1ul, "Me", 100.0, 200.0), (model.Markers[0].SteamId, model.Markers[0].OriginalName, model.Markers[0].X, model.Markers[0].Y));

        map.Markers = [new(9, MapMarkerKind.Player, 300, 400) { SteamId = 2 }];
        await model.ObserveAsync([M("1", "Me", true), M("2", "Bob", true)]);
        await model.ObserveAsync([M("1", "Me", false, 150, 250), M("2", "Bob", false)]);
        Assert.AreEqual(2, model.Markers.Count, "My cap is 1: the older pin goes.");
        Assert.AreEqual(150, model.Markers.Single(m => m.SteamId == 1).X);
        Assert.AreEqual((300.0, 400.0), (model.Markers.Single(m => m.SteamId == 2).X, model.Markers.Single(m => m.SteamId == 2).Y), "Bob's pin uses his player marker.");
        Assert.AreEqual(3, store.Saved.Count, "One save per pin.");

        team.Reset();
        await model.ObserveAsync([M("2", "Bob", false, 1, 1)]);
        Assert.AreEqual(2, model.Markers.Count, "After a reconnect the first read is only a sighting.");

        Assert.IsFalse(model.ShowWipeOverlay);
        model.Show = true;
        Assert.AreEqual(true, settings.Values[SettingKeys.MapShowDeathTags]);
        Assert.IsTrue(model.ShowWipeOverlay);
        map.AbbreviateNames = true;
        CollectionAssert.AreEquivalent(new[] { "M...", "B..." }, model.Pins.Select(pin => pin.Marker.DisplayName).ToArray(), "Streamer mode abbreviates pin names.");
        map.AbbreviateNames = false;

        var bob = model.Markers.Single(m => m.SteamId == 2);
        await model.RenameAsync(bob.Id, "Bob at the bridge");
        Assert.AreEqual("Bob at the bridge", model.Markers.Single(m => m.SteamId == 2).DisplayName);
        await model.RenameAsync(bob.Id, "  ");
        Assert.AreEqual("Bob", model.Markers.Single(m => m.SteamId == 2).DisplayName, "An empty name restores the player's name.");
        await model.DeleteAsync(bob.Id);
        Assert.AreEqual(1, model.Markers.Count);

        model.MaxTeam = 80;
        Assert.AreEqual(50, model.MaxTeam);
        Assert.AreEqual(50, settings.Values[SettingKeys.MaxTeamDeathMarkers]);
        store.Markers = [.. store.Markers, D(2, 1), D(2, 2), D(2, 3)];
        model.MaxTeam = 2;
        await model.ApplyCapsCommand.ExecuteAsync(null);
        Assert.AreEqual(2, model.Markers.Count(m => m.SteamId == 2), "Apply limits trims to the caps now.");

        await model.ConfirmWipeCommand.ExecuteAsync(null);
        Assert.IsTrue(model.IsWipeConfirming);
        Assert.AreEqual(3, model.Markers.Count, "The first click only asks.");
        model.CancelWipeCommand.Execute(null);
        Assert.IsFalse(model.IsWipeConfirming);
        await model.ConfirmWipeCommand.ExecuteAsync(null);
        await model.ConfirmWipeCommand.ExecuteAsync(null);
        Assert.AreEqual(0, model.Markers.Count);
        Assert.IsFalse(model.ShowWipeOverlay);

        store.Fail = true;
        await model.WipeCommand.ExecuteAsync(null);
        Assert.IsTrue(log.Contains("[deaths] The server profile changed; death markers were not saved."));
        return true;
    }, default);

    [TestMethod]
    public async Task Map_DrawsPins_HoversTheirLabel_AndTheWorkspaceMenuRenamesAndDeletes() => await UiTestSession.Current.Dispatch(async () =>
    {
        var store = new MemoryMarkers { Markers = [D(76561198000000001, 0, 1500, "Synthetic Builder"), D(76561198000000001, 3, 3200, "Synthetic Builder"), D(76561198000000002, 1, 800, "Farmer")] };
        var shell = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory(),
            appSettings: new ShopNoticeTests.MemorySettings(), deathMarkerStore: store);
        await shell.RefreshCommand.ExecuteAsync(null);
        await Task.Yield();
        Assert.AreEqual(3, shell.DeathPins!.Markers.Count, "Selecting the server loads its pins.");
        shell.Connection.Map.Apply(MapTests.SyntheticMap());
        shell.DeathPins.Show = true;
        var window = new MainWindow { DataContext = shell, Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark };
        try
        {
            window.Show();
            window.UpdateLayout();
            var surface = window.GetVisualDescendants().OfType<MapSurface>().Single();
            Assert.AreEqual(3, surface.DeathPins!.Count);
            var wipe = window.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "Wipe death markers on the map");
            Assert.IsTrue(wipe.IsEffectivelyVisible, "The map wipe button shows with pins.");

            using (var first = window.CaptureRenderedFrame()) Assert.IsNotNull(first); // the first render fits the map
            var pin = surface.DeathPins[0];
            var (ix, iy) = surface.Geometry!.WorldToImage(pin.Marker.X, pin.Marker.Y);
            var tip = surface.ImageToScreen(ix, iy);
            var head = new Point(tip.X, tip.Y - MapSurface.PinHeight + 8);
            Assert.AreSame(pin, surface.FindPin(head));
            Assert.IsNull(surface.FindPin(new Point(tip.X + 40, tip.Y + 40)));
            window.MouseMove(surface.TranslatePoint(head, window)!.Value);
            Assert.AreSame(pin, surface.HoveredPin);
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-death-pins.png"), PngBitmapEncoderOptions.Default);
            }

            var workspace = window.GetVisualDescendants().OfType<MapWorkspace>().Single();
            workspace.ShowPinMenu(pin);
            var items = workspace.PinMenu!.Items.OfType<MenuItem>().ToArray();
            CollectionAssert.AreEqual(new[] { "Rename Death Marker", "Delete Death Marker" }, items.Select(item => (string)item.Header!).ToArray());
            items[0].RaiseEvent(new global::Avalonia.Interactivity.RoutedEventArgs(MenuItem.ClickEvent));
            var box = (TextBox)((StackPanel)workspace.RenameFlyout!.Content!).Children[1];
            Assert.AreEqual("Synthetic Builder", box.Text);
            box.Text = "First base";
            var save = ((StackPanel)((StackPanel)workspace.RenameFlyout.Content!).Children[2]).Children.OfType<Button>().Single(b => (string)b.Content! == "Rename");
            save.RaiseEvent(new global::Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            await Task.Yield();
            Assert.AreEqual("First base", shell.DeathPins.Markers.Single(m => m.Id == pin.Marker.Id).CustomName);
            workspace.PinMenu.Close();

            workspace.ShowPinMenu(shell.DeathPins.Pins.Single(p => p.Marker.Id == pin.Marker.Id));
            workspace.PinMenu.Items.OfType<MenuItem>().Last().RaiseEvent(new global::Avalonia.Interactivity.RoutedEventArgs(MenuItem.ClickEvent));
            await Task.Yield();
            Assert.AreEqual(2, shell.DeathPins.Markers.Count);
            workspace.PinMenu.Close();

            wipe.Command!.Execute(null);
            await Task.Yield();
            Assert.AreEqual(0, shell.DeathPins.Markers.Count, "The map button wipes without asking, like the reference.");
        }
        finally { window.Close(); await shell.ShutdownAsync(); }
        return true;
    }, default);

    private sealed class MemoryMarkers : IDeathMarkerStore
    {
        public IReadOnlyList<DeathMarker> Markers { get; set; } = [];
        public List<IReadOnlyList<DeathMarker>> Saved { get; } = [];
        public bool Fail { get; set; }
        public Task<IReadOnlyList<DeathMarker>> LoadAsync(string profilesPath, string profileKey, CancellationToken ct) => Task.FromResult(Markers);
        public Task<IReadOnlyList<DeathMarker>> UpdateAsync(string profilesPath, string profileKey, Func<IReadOnlyList<DeathMarker>, IReadOnlyList<DeathMarker>> change,
            CancellationToken ct)
        {
            if (Fail) throw new ProfileChangedException();
            Markers = change(Markers);
            Saved.Add(Markers);
            return Task.FromResult(Markers);
        }
    }

    private sealed class OneServer : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([new("Rusty", "1.2.3.4", 28082) { ProfileKey = "key" }]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
