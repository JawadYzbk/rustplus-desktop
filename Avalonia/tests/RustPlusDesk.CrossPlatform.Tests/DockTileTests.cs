using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Connections;
using RustPlusDesk.Avalonia.Features.Dock;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Time;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MiniMapWindow.CommandDock: the tiles the dock draws - the device, the oil rig alarm, the logic rule, the
/// session, the death wipe and the map.
///
/// Every tile shows real data and does what it says: a tile whose data the app does not have is not offered at all,
/// and a tile that cannot act says why instead of doing nothing.
/// </summary>
[TestClass]
public sealed class DockTileTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-dock-tiles-" + Guid.NewGuid().ToString("N"));
        public LegacyCommandDockStore Store { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            Store = new LegacyCommandDockStore(Folder);
        }

        public void Dispose()
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
                    return;
                }
                catch (IOException)
                {
                    Thread.Sleep(50);
                }
            }
        }
    }

    /// <summary>A host that answers with whatever the test wants the dock to show.</summary>
    private sealed class FakeHost : IDockHost, IDockHostDevices, IDockHostContext
    {
        public string ServerTime { get; set; } = "15:20";
        public string? TimeUntilNextPhase { get; set; } = "12m until night";
        public string Population { get; set; } = "128 / 200";
        public string Queue { get; set; } = "3";
        public List<DockChatLine> Team { get; } = [];
        public List<DockChatLine> Clan { get; } = [];

        public IReadOnlyList<DockChatLine> Chat(bool clan, int count)
        {
            var feed = clan ? Clan : Team;
            return feed.Count <= count ? feed.ToList() : feed.GetRange(feed.Count - count, count);
        }

        // ---- the devices -------------------------------------------------------------------------

        public List<DockHostDevice> Devices { get; } = [];
        public IReadOnlyList<DockHostDevice> DockDevices => Devices;
        public List<(uint EntityId, bool On)> Switches { get; } = [];

        public Task<bool> SetDockDeviceAsync(uint entityId, bool on, CancellationToken ct)
        {
            Switches.Add((entityId, on));
            return Task.FromResult(true);
        }

        // ---- the rest of the app's state ---------------------------------------------------------

        public List<DockHostRule> Rules { get; } = [];
        public IReadOnlyList<DockHostRule> DockRules => Rules;
        public bool IsDockLogicEngineActive { get; set; } = true;
        public List<string> RunRules { get; } = [];

        public void RunDockRule(string ruleId) => RunRules.Add(ruleId);

        public DockHostEvent? LatestOilRigAlarm { get; set; }
        public string? DockServerKey { get; set; } = "10.0.0.1-28082";
        public SessionState? DockSession { get; set; }
        public int DockDeathMarkerCount { get; set; }
        public List<bool> Wipes { get; } = [];

        public Task WipeDockDeathMarkersAsync(bool keepLatest, CancellationToken ct)
        {
            Wipes.Add(keepLatest);
            DockDeathMarkerCount = 0;
            return Task.CompletedTask;
        }

        public MapViewModel? Map { get; set; }
        public IReadOnlyList<RustPlusDesk.Core.Deaths.DeathPin>? DeathPins { get; set; }
        public bool ShowDeathPins { get; set; }
        public bool MapOpened { get; private set; }

        public void OpenMap() => MapOpened = true;
    }

    private static async Task<DockViewModel> ALoadedDockAsync(Fixture fixture, FakeHost host)
    {
        var model = new DockViewModel(host, fixture.Store);
        await model.LoadAsync(CancellationToken.None);
        return model;
    }

    /// <summary>The tile of a kind, as it is drawn on the dock.</summary>
    private static DockTileView TheTile(Window window, string kind) =>
        window.GetVisualDescendants().OfType<DockTileView>().First(tile => tile.Tile.Kind == kind);

    /// <summary>The tile's own frame, whose opacity and border the tile changes.</summary>
    private static Border TheShell(DockTileView tile) =>
        tile.GetVisualDescendants().OfType<Border>().First(border => border.Name == "Shell");

    /// <summary>The texts the tile draws.</summary>
    private static List<string> TheTexts(DockTileView tile) =>
        tile.GetVisualDescendants().OfType<TextBlock>().Select(text => text.Text ?? "").ToList();

    /// <summary>The tile's own tip, the text the pointer rests on.</summary>
    private static string TheTip(DockTileView tile) => ToolTip.GetTip(TheShell(tile)) as string ?? "";

    private static Bitmap AMapImage()
    {
        var target = new RenderTargetBitmap(new PixelSize(256, 256));
        using (var context = target.CreateDrawingContext())
            context.FillRectangle(new SolidColorBrush(Color.FromRgb(18, 64, 82)), new Rect(0, 0, 256, 256));
        return target;
    }

    // ---- the device tile -------------------------------------------------------------------------

    [TestMethod]
    public async Task TheDeviceTileShowsTheSwitchAndTurnsItOnAndOff() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        host.Devices.Add(new DockHostDevice(42, "Workshop lamp", "SmartSwitch") { IsOn = true });
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 42 });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Device);
            Assert.IsTrue(TheTexts(tile).Contains("ON"), "The switch is on, and the tile says so.");
            Assert.IsTrue(TheTexts(tile).Contains("Workshop lamp"), "The device's name is on the tile.");

            var picture = tile.GetVisualDescendants().OfType<RustPlusDesk.Avalonia.Shared.Controls.ItemIcon>().FirstOrDefault();
            Assert.IsNotNull(picture, "The device's own picture is on the tile.");
            Assert.AreEqual("smart.switch", picture!.DefaultShortName, "A switch with no picture of its own falls back to the list's.");

            tile.Click();
            Assert.AreEqual(1, host.Switches.Count, "The click reached the device.");
            Assert.AreEqual((42u, false), host.Switches[0], "A switch that is on is turned off.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheDeviceTileOfADeviceThatIsGoneIsDimmed() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 99 });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Device);
            Assert.IsTrue(TheTexts(tile).Contains("not paired"), "The tile says the device is not paired.");
            Assert.AreEqual(0.5, TheShell(tile).Opacity, 0.01, "A device that is gone is dimmed.");

            tile.Click();
            Assert.AreEqual(0, host.Switches.Count, "There is nothing to switch.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheDeviceTileShowsTheAlarmAndTheUpkeep() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        host.Devices.Add(new DockHostDevice(7, "Raid alarm", "SmartAlarm") { IsOn = true });
        host.Devices.Add(new DockHostDevice(8, "Tool cupboard", "StorageMonitor") { UpkeepSeconds = 3900 });
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 7 });
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 8 });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var alarm = TheTile(window, CommandDockTileKinds.Device);
            Assert.IsTrue(TheTexts(alarm).Contains("ACTIVE"), "The alarm is going off, and the tile says so.");

            var cupboard = window.GetVisualDescendants().OfType<DockTileView>()
                .First(tile => tile.Tile.EntityId == 8);
            Assert.IsTrue(TheTexts(cupboard).Contains("1h 5m"), "The upkeep is on the tile.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the event tile --------------------------------------------------------------------------

    [TestMethod]
    public async Task TheEventTileShowsTheOilRigAlarmAndDimsAnOldOne() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost
        {
            LatestOilRigAlarm = new DockHostEvent(DateTime.UtcNow.AddMinutes(-2), "Small Oil Rig"),
        };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Event, EventKey = "oilrig" });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Event);
            Assert.IsTrue(TheTexts(tile).Any(text => text.Contains(":")), "The time of the alarm is on the tile.");
            Assert.AreEqual(1.0, TheShell(tile).Opacity, 0.01, "An alarm under fifteen minutes old is live.");

            // An alarm that is more than fifteen minutes old is dimmed.
            host.LatestOilRigAlarm = new DockHostEvent(DateTime.UtcNow.AddMinutes(-20), "Small Oil Rig");
            model.Refresh();
            Assert.AreEqual(0.5, TheShell(tile).Opacity, 0.01, "An old alarm is dimmed.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheEventTileSaysThatNoAlarmHasGoneOffYet() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Event, EventKey = "oilrig" });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Event);
            Assert.IsTrue(TheTexts(tile).Contains("—"), "There is no alarm to show yet.");
            Assert.AreEqual(0.5, TheShell(tile).Opacity, 0.01, "An alarm that has not gone off is dimmed.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the rule tile ---------------------------------------------------------------------------

    [TestMethod]
    public async Task TheRuleTileRunsItsRule() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        host.Rules.Add(new DockHostRule("rule-1", "Raid the doors", true, DockRuleTriggers.CommandDock));
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Rule, RuleId = "rule-1" });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Rule);
            Assert.IsTrue(TheTexts(tile).Contains("Raid the doors"), "The rule's name is on the tile.");

            tile.Click();
            CollectionAssert.AreEqual(new[] { "rule-1" }, host.RunRules.ToArray(), "The click ran the rule.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheRuleTileSaysWhyItCannotRun() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        host.Rules.Add(new DockHostRule("rule-1", "Raid the doors", true, "SmartAlarm"));
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Rule, RuleId = "rule-1" });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Rule);
            StringAssert.Contains(TheTip(tile), "Widget Overlay");
            Assert.AreEqual(0.4, TheShell(tile).Opacity, 0.01, "A rule that cannot run is dimmed.");

            tile.Click();
            Assert.AreEqual(0, host.RunRules.Count, "A rule that cannot run does not run.");

            // A rule that is switched off says so as well.
            host.Rules[0] = new DockHostRule("rule-1", "Raid the doors", false, DockRuleTriggers.CommandDock);
            model.Refresh();
            StringAssert.Contains(TheTip(tile), "Activate");

            // And a rule that is gone says that too.
            host.Rules.Clear();
            model.Refresh();
            StringAssert.Contains(TheTip(tile), "no longer exists");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the session tile ------------------------------------------------------------------------

    [TestMethod]
    public async Task TheSessionTileShowsTheFigures() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost
        {
            DockSession = new SessionState
            {
                ServerKey = "10.0.0.1-28082",
                OnlineSeconds = 4500,
                AfkSeconds = 900,
                DistanceMetres = 1500,
                Deaths = 2,
                TeamDeaths = 1,
            },
        };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Session, ColSpan = 3 });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var texts = TheTexts(TheTile(window, CommandDockTileKinds.Session));
            Assert.IsTrue(texts.Contains("1:15h"), "The time online is on the tile.");
            Assert.IsTrue(texts.Contains("15m"), "The idle time is on the tile.");
            Assert.IsTrue(texts.Contains("1.5 km"), "The distance walked is on the tile.");
            Assert.IsTrue(texts.Contains("2"), "The deaths are on the tile.");
            Assert.IsTrue(texts.Contains("1"), "The team's deaths are on the tile.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheSessionTileSaysThatItIsWaitingForTheServer() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost { DockSession = null };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Session, ColSpan = 3 });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Session);
            Assert.IsTrue(TheTexts(tile).All(text => text == "—" || text.Length > 0), "The figures are placeholders.");
            StringAssert.Contains(TheTip(tile), "connected");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the death wipe tile ----------------------------------------------------------------------

    [TestMethod]
    public async Task TheDeathWipeTileCountsTheMarkersAndClearsThem() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost { DockDeathMarkerCount = 3 };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.DeathWipe, DeathWipeAlwaysVisible = true });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.DeathWipe);
            Assert.IsTrue(TheTexts(tile).Contains("3"), "The tile counts the markers on the map.");

            tile.Click();
            CollectionAssert.AreEqual(new[] { false }, host.Wipes.ToArray(), "The press cleared every marker.");
            Assert.AreEqual(0, host.DockDeathMarkerCount, "The markers are gone.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheDeathWipeTileIsDimmedWhenThereIsNothingToClear() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost { DockDeathMarkerCount = 0 };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.DeathWipe, DeathWipeAlwaysVisible = true });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.DeathWipe);
            Assert.AreEqual(0.6, TheShell(tile).Opacity, 0.01, "A wipe with nothing to do is dimmed.");
            StringAssert.Contains(TheTip(tile), "No death markers");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheDeathWipeTileGivesItsCellsBackUntilThereIsSomethingToClear() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost { DockDeathMarkerCount = 0 };
        var model = await ALoadedDockAsync(fixture, host);
        var tile = model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.DeathWipe });

        Assert.IsFalse(model.BelongsHere(tile), "Nothing to clear, so the tile is hidden where it stands.");

        tile.DeathWipeAlwaysVisible = true;
        Assert.IsTrue(model.BelongsHere(tile), "A tile set to stay keeps its place.");

        host.DockDeathMarkerCount = 2;
        tile.DeathWipeAlwaysVisible = false;
        Assert.IsTrue(model.BelongsHere(tile), "There is something to clear, so the tile comes back.");
        return true;
    }, default);

    // ---- the map tile -----------------------------------------------------------------------------

    [TestMethod]
    public async Task TheMapTileDrawsTheMapAndOpensTheWorkspace() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost
        {
            Map = new MapViewModel { Image = AMapImage(), Geometry = new MapGeometry(256, 256, 4500) },
        };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Map, Id = CommandDockTileKinds.MapTileId });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = TheTile(window, CommandDockTileKinds.Map);
            var surface = tile.GetVisualDescendants().OfType<MapSurface>().First();
            Assert.IsNotNull(surface.Image, "The tile draws the map the workspace draws.");
            Assert.AreSame(host.Map!.Image, surface.Image, "It is the same picture.");

            tile.Click();
            Assert.IsTrue(host.MapOpened, "A click on the map opens the map workspace.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the whole dock ----------------------------------------------------------------------------

    [TestMethod]
    public async Task TheDockDrawsEveryKindThisSliceAdds() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost { DockDeathMarkerCount = 1 };
        host.Devices.Add(new DockHostDevice(42, "Workshop lamp", "SmartSwitch") { IsOn = false });
        host.Rules.Add(new DockHostRule("rule-1", "Raid the doors", true, DockRuleTriggers.CommandDock));
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 42 });
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Event, EventKey = "oilrig" });
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Rule, RuleId = "rule-1" });
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.Session, ColSpan = 3 });
        model.AddTile(new CommandDockTile { Kind = CommandDockTileKinds.DeathWipe });

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var kinds = window.GetVisualDescendants().OfType<DockTileView>().Select(tile => tile.Tile.Kind).ToArray();
            foreach (var kind in new[]
                     {
                         CommandDockTileKinds.Map, CommandDockTileKinds.Device, CommandDockTileKinds.Event,
                         CommandDockTileKinds.Rule, CommandDockTileKinds.Session, CommandDockTileKinds.DeathWipe,
                     })
            {
                Assert.IsTrue(kinds.Contains(kind), $"The dock carries a {kind} tile.");
            }
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public void ADeviceTileOfAnotherServerIsHiddenWhereItStands()
    {
        using var fixture = new Fixture();
        var host = new FakeHost { DockServerKey = "10.0.0.2-28082" };
        var model = new DockViewModel(host, fixture.Store);

        var mine = new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 42, ServerKey = "10.0.0.2-28082" };
        var theirs = new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 43, ServerKey = "10.0.0.1-28082" };

        Assert.IsTrue(model.BelongsHere(mine), "A device tile of the connected server stays.");
        Assert.IsFalse(model.BelongsHere(theirs), "An entity id means nothing on another server.");
    }

    // ---- the catalogue ----------------------------------------------------------------------------
    private static IReadOnlyList<DockTileOption> Options(bool mapPresent, FakeHost host)
    {
        var options = new List<DockTileOption>();
        foreach (var section in DockTileCatalogue.Sections(mapPresent, host))
            options.AddRange(section.Options);
        return options;
    }

    [TestMethod]
    public void ThePickerOffersTheKindsThisSliceDraws()
    {
        var host = new FakeHost();
        host.Devices.Add(new DockHostDevice(42, "Workshop lamp", "SmartSwitch"));
        host.Rules.Add(new DockHostRule("rule-1", "Raid the doors", true, DockRuleTriggers.CommandDock));

        var options = Options(mapPresent: true, host);

        Assert.IsTrue(options.Any(option => option.Kind == CommandDockTileKinds.Device), "A device is offered.");
        Assert.IsTrue(options.Any(option => option.Kind == CommandDockTileKinds.Event), "The oil rig alarm is offered.");
        Assert.IsTrue(options.Any(option => option.Kind == CommandDockTileKinds.Rule), "A rule is offered.");
        Assert.IsTrue(options.Any(option => option.Kind == CommandDockTileKinds.Session), "The session is offered.");
        Assert.IsTrue(options.Any(option => option.Kind == CommandDockTileKinds.DeathWipe), "The death wipe is offered.");

        foreach (var option in options)
            Assert.IsTrue(DockTiles.IsSupported(option.Kind), $"The picker offers {option.Kind}, which this build cannot draw.");
    }

    [TestMethod]
    public void ThePickerOffersOneEntryPerPairedDeviceAndStampsItWithItsServer()
    {
        var host = new FakeHost { DockServerKey = "10.0.0.1-28082" };
        host.Devices.Add(new DockHostDevice(42, "Workshop lamp", "SmartSwitch"));
        host.Devices.Add(new DockHostDevice(43, "Raid alarm", "SmartAlarm"));

        var devices = Options(mapPresent: true, host).Where(option => option.Kind == CommandDockTileKinds.Device).ToArray();

        Assert.AreEqual(2, devices.Length, "One entry for each paired device.");
        foreach (var device in devices)
        {
            var tile = device.Make();
            Assert.AreEqual("10.0.0.1-28082", tile.ServerKey, "The tile is stamped with the server it belongs to.");
        }
    }

    [TestMethod]
    public void ThePickerOffersOnlyTheRulesThatAskedToBeLaunchedByHand()
    {
        var host = new FakeHost();
        host.Rules.Add(new DockHostRule("rule-1", "Raid the doors", true, DockRuleTriggers.CommandDock));
        host.Rules.Add(new DockHostRule("rule-2", "Countdown", true, "SmartAlarm"));

        var rules = Options(mapPresent: true, host).Where(option => option.Kind == CommandDockTileKinds.Rule).ToArray();

        Assert.AreEqual(1, rules.Length, "Only the rule that asked for the dock's own trigger is offered.");
        Assert.AreEqual("rule-1", rules[0].Make().RuleId, "It is the right one.");
    }

    [TestMethod]
    public void ATileWhoseDataTheAppDoesNotHaveIsNotDrawn()
    {
        var tile = new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = 1 };

        Assert.IsTrue(DockTiles.IsSupported(CommandDockTileKinds.Device), "The build draws device tiles.");
        Assert.IsFalse(DockTiles.CanServe(tile, new BareHost()), "A host that cannot answer the device is not served.");
        Assert.IsTrue(DockTiles.CanServe(tile, new FakeHost()), "A host that can answer it is.");
    }

    /// <summary>A host that answers nothing but the clock: what a connection with no app behind it looks like.</summary>
    private sealed class BareHost : IDockHost
    {
        public string ServerTime => "15:20";
        public string? TimeUntilNextPhase => null;
        public string Population => "0 / 0";
        public string Queue => "";
        public IReadOnlyList<DockChatLine> Chat(bool clan, int count) => [];
    }

    // ---- the session tracker's wiring --------------------------------------------------------------

    [TestMethod]
    public void TheSessionTrackerCountsWhatTheConnectionSays()
    {
        var connection = new ConnectionViewModel(new UnusedSessionFactory());
        connection.Phase = ConnectionPhase.Connected;
        connection.Population = "128 / 200";
        connection.Team.Members =
        [
            new TeamMemberRow(new ServerTeamMember("76561198000000001", "Me", true, true) { X = 10, Y = 20 }, AfkText: "5:00"),
            new TeamMemberRow(new ServerTeamMember("76561198000000002", "Them", true, true) { X = 30, Y = 40 }),
        ];

        var source = new DockSessionSource(connection,
            () => "10.0.0.1-28082",
            () => 76561198000000001UL);

        var snapshot = source.Read();
        Assert.IsNotNull(snapshot, "A connected server is read.");
        Assert.IsTrue(snapshot!.Connected);
        Assert.AreEqual("10.0.0.1-28082", snapshot.ServerKey);
        Assert.AreEqual(76561198000000001UL, snapshot.MySteamId);
        Assert.AreEqual(10, snapshot.MyX);
        Assert.AreEqual(20, snapshot.MyY);
        Assert.IsTrue(snapshot.MyAfk, "The player is away from the keyboard.");
        Assert.AreEqual(128, snapshot.Population);
        Assert.AreEqual(2, snapshot.Team.Count);

        // The tracker counts what it is handed.
        var clock = new FixedClock(new DateTimeOffset(2026, 10, 20, 12, 0, 0, TimeSpan.Zero));
        var timer = new ManualTimer();
        using var tracker = new SessionTracker(clock, timer, new MemoryStore());
        tracker.Source = () => source.Read();

        clock.Now = clock.Now.AddSeconds(1);
        tracker.Tick();

        // The player walks ten metres between the two readings.
        connection.Team.Members =
        [
            new TeamMemberRow(new ServerTeamMember("76561198000000001", "Me", true, true) { X = 20, Y = 20 }, AfkText: "5:00"),
            new TeamMemberRow(new ServerTeamMember("76561198000000002", "Them", true, true) { X = 30, Y = 40 }),
        ];
        clock.Now = clock.Now.AddSeconds(1);
        tracker.Tick();

        var session = tracker.For("10.0.0.1-28082");
        Assert.IsNotNull(session, "The session of the connected server is kept.");
        Assert.AreEqual(2, session!.OnlineSeconds, 0.001, "Two ticks are two seconds online.");
        Assert.AreEqual(2, session.AfkSeconds, 0.001, "The player was away the whole time.");
        Assert.AreEqual(10, session.DistanceMetres, 0.001, "The distance walked is counted.");
    }

    /// <summary>A timer the test fires by hand.</summary>
    private sealed class ManualTimer : ISessionTimer
    {
        public TimeSpan Interval { get; set; } = TimeSpan.FromSeconds(1);
        public event EventHandler? Ticking;
        public void Start() { }
        public void Stop() { }
        public void Fire() => Ticking?.Invoke(this, EventArgs.Empty);
    }

    private sealed class MemoryStore : ISessionTrackerStore
    {
        public Dictionary<string, SessionState> Saved { get; } = new(StringComparer.Ordinal);
        public SessionState? Load(string serverKey) => Saved.TryGetValue(serverKey, out SessionState? state) ? state : null;
        public void Save(SessionState state) => Saved[state.ServerKey] = state;
    }
}
