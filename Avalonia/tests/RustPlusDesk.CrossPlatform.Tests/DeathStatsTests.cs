using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Deaths;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Deaths;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeathStatsTests
{
    // Wall-clock based: the stats page filters its ranges against DateTimeOffset.Now, so the sample timestamps must
    // follow the same clock or the 24 h/7 d assertions drift out of the window the day after they are written.
    private static readonly DateTimeOffset Now = DateTimeOffset.Now;
    private static long Ago(double hours) => Now.AddHours(-hours).ToUnixTimeSeconds();

    private static string Line(string name, long died, long? spawn, string grid, string type, string? place) =>
        System.Text.Json.JsonSerializer.Serialize(new { steam_id = "1", name, died_at = died, spawn_at = spawn, x = 1.0, y = 2.0, grid, location_type = type, location_name = place });

    private static string[] SampleLog() =>
    [
        Line("Ann", Ago(1), Ago(1) - 600, "G7", "monument", "Harbor"),
        Line("Ann", Ago(2), Ago(2) - 30, "G7", "base", "Main"),
        Line("Bob", Ago(30), null, "off-grid", "open", null),
        Line("Bob", Ago(200), Ago(200) - 7200, "A1", "monument", "Harbor"),
        "not json",
        ""
    ];

    [TestMethod]
    public void Summary_FollowsTheReferenceAggregation()
    {
        var entries = DeathStats.ParseEntries(SampleLog());
        Assert.AreEqual(4, entries.Count, "Malformed and blank lines are skipped.");
        Assert.AreEqual("Open", entries[2].Location, "An unnamed place is labelled by its area.");
        var summary = DeathStats.Summarize(entries, null, false, Now.LocalDateTime.Date);
        Assert.AreEqual(4, summary.Total);
        Assert.AreEqual(2, summary.Victims);
        Assert.AreEqual("4 death(s) across 2 player(s).", summary.Headline);
        Assert.AreEqual("Harbor", summary.DeadliestPlace);
        Assert.AreEqual("G7 (2)", summary.DeadliestGrid, "off-grid never wins the deadliest grid.");
        Assert.AreEqual("2h 0m", summary.LongestSurvival);
        Assert.AreEqual(DeathStats.FormatDuration((600 + 30 + 7200) / 3), summary.AvgSurvival);
        CollectionAssert.AreEqual(new[] { "Monument", "Base", "Open" }, summary.ByArea.Select(a => a.Name).ToArray());
        CollectionAssert.AreEqual(new[] { 50, 25, 25 }, summary.ByArea.Select(a => a.Percent).ToArray());
        Assert.AreEqual(14, summary.DeathsPerDay.Count);
        Assert.AreEqual(Now.LocalDateTime.Date, summary.DeathsPerDay[^1].Day);
        Assert.AreEqual(4, summary.Recent.Count);
        Assert.AreEqual("—", summary.Recent[2].Grid == "off-grid" ? "—" : summary.Recent[2].Grid);
        Assert.AreEqual("No deaths logged on this server yet. They appear here as they happen.", DeathStats.Summarize([], null, true, Now.Date).Headline);
        Assert.AreEqual("No deaths match the current filters.", DeathStats.Summarize([], null, false, Now.Date).Headline);
        Assert.AreEqual("45s", DeathStats.FormatDuration(45));
        Assert.AreEqual("2m 5s", DeathStats.FormatDuration(125));
    }

    [TestMethod]
    public void Filters_SearchPlayerPlaceAndRange()
    {
        var entries = DeathStats.ParseEntries(SampleLog());
        long now = Now.ToUnixTimeSeconds();
        Assert.AreEqual(2, DeathStats.Filter(entries, "harbor", "all", "all", "all", now).Count());
        Assert.AreEqual(2, DeathStats.Filter(entries, " g7 ", "all", "all", "all", now).Count(), "Grid search, trimmed, ignoring case.");
        Assert.AreEqual(2, DeathStats.Filter(entries, "", "Bob", "all", "all", now).Count());
        Assert.AreEqual(1, DeathStats.Filter(entries, "", "all", "base", "all", now).Count());
        Assert.AreEqual(2, DeathStats.Filter(entries, "", "all", "all", "24h", now).Count());
        Assert.AreEqual(3, DeathStats.Filter(entries, "", "all", "all", "7d", now).Count());
    }

    [TestMethod]
    public void Killers_LastLineWins_GroupIgnoringCase_AndForgetKeepsUnreadableLines()
    {
        string[] lines =
        [
            "{\"died_at\":10,\"killer\":\"Raider\",\"weapon\":\"AK\",\"recorded_at\":1}",
            "{\"died_at\":20,\"killer\":\"raider\",\"weapon\":\"Bow\",\"recorded_at\":2}",
            "{\"died_at\":30,\"killer\":\"Wrong\",\"weapon\":null,\"recorded_at\":3}",
            "{\"died_at\":30,\"killer\":\"Other\",\"weapon\":null,\"recorded_at\":4}",
            "garbage"
        ];
        var byDeath = DeathStats.ParseKillers(lines);
        Assert.AreEqual(("Other", (string?)null), byDeath[30], "A correction written later wins.");
        var stats = DeathStats.SummarizeKillers(byDeath);
        Assert.AreEqual("raider", stats[0].Name, "The latest spelling is kept.");
        Assert.AreEqual(2, stats[0].Kills);
        Assert.AreEqual("Bow", stats[0].LastWeapon);
        Assert.AreEqual(1, DeathStats.SummarizeKillers(byDeath, [10, 999]).Single().Kills, "Only the filtered deaths count.");
        CollectionAssert.AreEqual(new[] { lines[2], lines[3], "garbage" }, lines.Where(line => DeathStats.KeepKillerLine(line, " RAIDER ")).ToArray());
    }

    [TestMethod]
    public async Task Page_OpensForTheSelectedServer_FiltersForgetsAndClears() => await UiTestSession.Current.Dispatch(async () =>
    {
        string root = Path.Combine(Path.GetTempPath(), "rpd-stats-" + Guid.NewGuid().ToString("N"));
        try
        {
            var log = new DeathLogFile(root);
            Directory.CreateDirectory(root);
            File.WriteAllLines(log.PathFor("1.2.3.4-28082"), SampleLog());
            File.WriteAllLines(log.KillerPathFor("1.2.3.4-28082"),
                ["{\"died_at\":" + Ago(1) + ",\"killer\":\"Raider\",\"weapon\":\"AK\",\"recorded_at\":1}",
                 "{\"died_at\":" + Ago(2) + ",\"killer\":\"Misread\",\"weapon\":null,\"recorded_at\":1}"]);
            var shell = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory(), deathLog: log);
            await shell.RefreshCommand.ExecuteAsync(null);
            Assert.IsTrue(shell.HasDeathStats);
            shell.OpenDeathStatsCommand.Execute(null);
            Assert.IsTrue(shell.IsDeathStatsSelected && shell.IsToolSelected);
            var page = shell.DeathStats!;
            Assert.AreEqual(4, page.Summary.Total);
            CollectionAssert.AreEqual(new[] { "All players", "Ann", "Bob" }, page.Players.Select(p => p.Label).ToArray());
            Assert.AreEqual(14, page.Sparkline.Count);
            Assert.IsTrue(page.Sparkline.All(bar => bar.Height is >= 2 and <= DeathStatsViewModel.SparkHeight));
            Assert.AreEqual(2, page.Summary.ByKiller.Count);

            var window = new RustPlusDesk.Avalonia.Shell.MainWindow { DataContext = shell, Width = 1000, Height = 700 };
            try
            {
                window.Show();
                window.UpdateLayout();
                Assert.IsNotNull(window.GetVisualDescendants().OfType<Button>().SingleOrDefault(b => AutomationProperties.GetName(b) == "Death Stats"), "Rail entry");
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-death-stats.png"), PngBitmapEncoderOptions.Default);
                }
                shell.IsSidebarOpen = false;
                window.UpdateLayout();
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-death-stats-wide.png"), PngBitmapEncoderOptions.Default);
                }
            }
            finally { window.Close(); }

            page.Player = page.Players.Single(p => p.Tag == "Bob");
            Assert.AreEqual(2, page.Summary.Total);
            Assert.AreEqual(0, page.Summary.ByKiller.Count, "Killers follow the filters.");
            page.Player = page.Players[0];
            page.Range = DeathStatsViewModel.RangeOptions.Single(o => o.Tag == "24h");
            Assert.AreEqual(2, page.Summary.Total);
            page.SearchText = "nobody";
            Assert.AreEqual("No deaths match the current filters.", page.Summary.Headline);
            page.SearchText = "";
            page.Range = DeathStatsViewModel.RangeOptions[0];

            page.ForgetKillerCommand.Execute("misread");
            Assert.AreEqual("Raider", page.Summary.ByKiller.Single().Name);
            page.RequestClearKillersCommand.Execute(null);
            Assert.IsTrue(page.IsConfirmingClearKillers);
            page.ConfirmClearKillersCommand.Execute(null);
            Assert.AreEqual(0, page.Summary.ByKiller.Count);
            Assert.IsFalse(File.Exists(log.KillerPathFor("1.2.3.4-28082")));
            Assert.AreEqual(4, page.Summary.Total, "Clearing killers keeps the deaths.");

            File.AppendAllLines(log.PathFor("1.2.3.4-28082"), [Line("Cid", Ago(0.5), null, "B2", "open", null)]);
            page.ReloadCommand.Execute(null);
            Assert.AreEqual(5, page.Summary.Total, "Refresh re-reads the log.");
            page.RequestClearCommand.Execute(null);
            page.CancelClearCommand.Execute(null);
            Assert.AreEqual(5, page.Summary.Total);
            page.RequestClearCommand.Execute(null);
            page.ConfirmClearCommand.Execute(null);
            Assert.IsFalse(File.Exists(log.PathFor("1.2.3.4-28082")));
            Assert.AreEqual("No deaths logged on this server yet. They appear here as they happen.", page.Summary.Headline);
        }
        finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        return true;
    }, default);

    [TestMethod]
    public async Task DeathHeatmap_LoadsTheSelectedServersPositionsWhileOn_AndPersists() => await UiTestSession.Current.Dispatch(async () =>
    {
        string root = Path.Combine(Path.GetTempPath(), "rpd-heat-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(root);
            var log = new DeathLogFile(root);
            File.WriteAllLines(log.PathFor("1.2.3.4-28082"), SampleLog());
            var settings = new ShopNoticeTests.MemorySettings();
            var shell = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory(), appSettings: settings, deathLog: log);
            await shell.RefreshCommand.ExecuteAsync(null);
            Assert.AreEqual(0, shell.Connection.Map.DeathPositions.Count, "Off by default: nothing is read.");
            shell.Connection.Map.ShowDeathHeatmap = true;
            Assert.AreEqual(true, settings.Values[RustPlusDesk.Core.Settings.SettingKeys.MapShowDeathHeatmap]);
            Assert.AreEqual(4, shell.Connection.Map.DeathPositions.Count);
            shell.Deaths!.Observe(new([new("7", "Eve", true, true)], [], []));
            shell.Deaths.Observe(new([new("7", "Eve", true, false) { DeathTime = 99, X = 5, Y = 6 }], [], []));
            Assert.AreEqual(5, shell.Connection.Map.DeathPositions.Count, "A new death shows up live.");
            Assert.AreEqual(new RustPlusDesk.Avalonia.Features.Map.Point2(5, 6), shell.Connection.Map.DeathPositions[^1]);
            shell.Connection.Map.ShowDeathHeatmap = false;
            Assert.AreEqual(0, shell.Connection.Map.DeathPositions.Count);

            var map = new RustPlusDesk.Avalonia.Features.Map.MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            var surface = new RustPlusDesk.Avalonia.Features.Map.MapSurface { Image = map.Image, Geometry = map.Geometry, ShowDeathHeatmap = true,
                DeathPositions = [new(2200, 2300), new(2230, 2300), new(2260, 2330), new(1200, 3200)] };
            var window = new Window { Width = 900, Height = 700, Content = surface };
            try
            {
                window.Show();
                using var frame = window.CaptureRenderedFrame();
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-death-heatmap.png"), PngBitmapEncoderOptions.Default);
            }
            finally { window.Close(); }
        }
        finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        return true;
    }, default);

    [TestMethod]
    public async Task MapWipe_ClearsTheSelectedServersDeathLog() => await UiTestSession.Current.Dispatch(async () =>
    {
        string root = Path.Combine(Path.GetTempPath(), "rpd-wipe-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(root);
            var log = new DeathLogFile(root);
            File.WriteAllLines(log.PathFor("1.2.3.4-28082"), SampleLog());
            File.WriteAllLines(log.PathFor("5.6.7.8-28082"), SampleLog());
            var cache = new RustPlusDesk.Infrastructure.Map.MapCacheStore(Path.Combine(root, "map_cache"));
            var map = MapTests.SyntheticMap();
            var old = new DateTimeOffset(2026, 9, 1, 0, 0, 0, TimeSpan.Zero);
            cache.Save("1.2.3.4", 28082, new RustPlusDesk.Core.Map.CachedMap(map.JpgImage, 1024, 1024, 4500, old, map.Monuments));
            var shell = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory(),
                appSettings: new ShopNoticeTests.MemorySettings(), deathLog: log, mapCache: cache);
            await shell.RefreshCommand.ExecuteAsync(null);
            shell.Connection.Map.ShowDeathHeatmap = true;
            Assert.AreEqual(4, shell.Connection.Map.DeathPositions.Count);

            await shell.Connection.Map.Loader!(new WipeSession(map, old.AddSeconds(20)), default);
            Assert.AreEqual(SampleLog().Length, log.ReadLines("1.2.3.4-28082").Count, "The same wipe keeps the log.");
            await shell.Connection.Map.Loader!(new WipeSession(map, old.AddDays(14)), default);
            Assert.AreEqual(0, log.ReadLines("1.2.3.4-28082").Count, "A new wipe clears the selected server's log.");
            Assert.AreEqual(SampleLog().Length, log.ReadLines("5.6.7.8-28082").Count, "Other servers keep theirs.");
            Assert.AreEqual(0, shell.Connection.Map.DeathPositions.Count, "The heatmap follows.");
            Assert.IsTrue(shell.Activity.Lines.Any(line => line.Contains("[deaths] Wipe detected; cleared the death log.", StringComparison.Ordinal)));
            Assert.AreEqual(old.AddDays(14), cache.Load("1.2.3.4", 28082)!.WipeTime, "The fresh map is cached for the new wipe.");
        }
        finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        return true;
    }, default);

    private sealed class WipeSession(RustPlusDesk.Core.Map.ServerMap map, DateTimeOffset wipe) : RustPlusDesk.Core.Map.IServerMapSession
    {
        public RustPlusDesk.Core.Devices.IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Synthetic", 0, 0, 0, WipeTime: wipe);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public Task<RustPlusDesk.Core.Map.ServerMap> ReadMapAsync(CancellationToken ct) => Task.FromResult(map);
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
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
