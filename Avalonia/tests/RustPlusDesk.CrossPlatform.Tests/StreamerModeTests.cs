using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Notifications;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Notifications;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class StreamerModeTests
{
    [TestMethod]
    public void Abbreviations_FollowTheReferenceRules()
    {
        Assert.AreEqual("J...", StreamerMode.Name("Jawad", true));
        Assert.AreEqual("Jawad", StreamerMode.Name("Jawad", false));
        Assert.AreEqual("  ", StreamerMode.Name("  ", true), "Blank names stay as they are.");
        Assert.AreEqual("765...", StreamerMode.SteamId("76561198000000001", true));
        Assert.AreEqual("12...", StreamerMode.SteamId("12", true));
        Assert.AreEqual("76561198000000001", StreamerMode.SteamId("76561198000000001", false));
    }

    [TestMethod]
    public void MapSettings_LoadAndPersistStreamerModeAndPlayerIconScale()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.MapAbbreviateNames, true);
        settings.Set(SettingKeys.MapPlayerIconScale, 9.0);
        var map = new MapViewModel();
        map.UseSettings(settings);
        Assert.IsTrue(map.AbbreviateNames);
        Assert.AreEqual(1.5, map.PlayerIconScale, "The saved scale is clamped to 50–150 %.");
        map.AbbreviateNames = false;
        map.PlayerIconScale = 0.75;
        Assert.AreEqual(false, settings.Values[SettingKeys.MapAbbreviateNames]);
        Assert.AreEqual(0.75, settings.Values[SettingKeys.MapPlayerIconScale]);
        map.PlayerIconScale = 0.1;
        Assert.AreEqual(0.5, map.PlayerIconScale);
        Assert.AreEqual(0.5, settings.Values[SettingKeys.MapPlayerIconScale]);
    }

    [TestMethod]
    public void TeamRows_AbbreviateNamesAndSteamIds()
    {
        var team = new TeamViewModel { Members = [new TeamMemberRow(new ServerTeamMember("76561198000000001", "Builder", true, true)),
            new TeamMemberRow(new ServerTeamMember("76561198000000002", "", false))] };
        team.Abbreviate = true;
        CollectionAssert.AreEqual(new[] { "B...", "765..." }, team.Members.Select(row => row.Name).ToArray());
        Assert.AreEqual("765...", team.Members[0].SteamId);
        team.Abbreviate = false;
        CollectionAssert.AreEqual(new[] { "Builder", "76561198000000002" }, team.Members.Select(row => row.Name).ToArray());
    }

    [TestMethod]
    public void Notifications_HideServerTags_AndSearchIgnoresServerNames()
    {
        var model = new NotificationsViewModel(null, null);
        model.Add(new InboxNotification { Type = "Alarm", Title = "Raid", Message = "Base", ServerIp = "1.2.3.4", ServerPort = 1, ServerName = "Rustoria" });
        Assert.AreEqual("Rustoria", model.Rows[0].ServerTag);
        model.IsStreamerMode = true;
        Assert.AreEqual(StreamerMode.HiddenServer, model.Rows[0].ServerTag);
        model.SearchText = "rustoria";
        Assert.AreEqual(0, model.Rows.Count, "Streamer mode: the search does not reveal server names.");
        model.SearchText = "raid";
        Assert.AreEqual(1, model.Rows.Count);
        model.RequestMuteCommand.Execute(model.Rows[0]);
        StringAssert.Contains(model.MuteQuestion, "from this server?");
        model.IsStreamerMode = false;
        model.SearchText = "rustoria";
        Assert.AreEqual(1, model.Rows.Count);
    }

    [TestMethod]
    public async Task Shell_OneSwitchCoversHeaderTeamInboxAndServerList()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var shell = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory(), appSettings: settings);
        await shell.RefreshCommand.ExecuteAsync(null);
        _ = shell.Notifications;
        Assert.AreEqual("Rustoria Main", shell.Connection.DisplayServerName);
        Assert.IsTrue(shell.IsSidebarOpen);

        shell.SettingsPage.Map!.AbbreviateNames = true;
        Assert.AreSame(shell.Connection.Map, shell.SettingsPage.Map, "Settings and the map toolbar share one switch.");
        Assert.AreEqual(StreamerMode.HiddenServer, shell.Connection.DisplayServerName);
        Assert.IsTrue(shell.Connection.Team.Abbreviate);
        Assert.IsTrue(shell.Notifications.IsStreamerMode);
        Assert.IsFalse(shell.IsSidebarOpen, "Turning streamer mode on closes the server list (reference).");
        Assert.AreEqual(true, settings.Values[SettingKeys.MapAbbreviateNames]);

        shell.Connection.Map.AbbreviateNames = false;
        Assert.AreEqual("Rustoria Main", shell.Connection.DisplayServerName);
        Assert.IsFalse(shell.Connection.Team.Abbreviate);
        Assert.IsFalse(shell.Notifications.IsStreamerMode);

        settings.Set(SettingKeys.MapAbbreviateNames, true);
        var restored = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new UnusedSessionFactory(), appSettings: settings);
        await restored.RefreshCommand.ExecuteAsync(null);
        Assert.AreEqual(StreamerMode.HiddenServer, restored.Connection.DisplayServerName, "A saved streamer mode applies on start.");
        Assert.IsTrue(restored.IsSidebarOpen, "Starting in streamer mode does not fold the sidebar by itself.");
        Assert.IsTrue(restored.Notifications.IsStreamerMode);
    }

    [TestMethod]
    public async Task MapSurface_AbbreviatesPlayerLabelsAndScalesMarkers() => await UiTestSession.Current.Dispatch(async () =>
    {
        var model = new MapViewModel();
        model.Apply(MapTests.SyntheticMap());
        var surface = new MapSurface { Image = model.Image, Geometry = model.Geometry, Monuments = model.Monuments, Markers = MapTests.SyntheticMarkers(),
            AbbreviateNames = true, PlayerScale = 1.5 };
        var window = new Window { Width = 900, Height = 700, Content = surface };
        try
        {
            window.Show();
            var player = surface.Markers!.First(marker => marker.Kind == MapMarkerKind.Player);
            StringAssert.StartsWith(MapSurface.Describe(player, 4500, abbreviateNames: true), "T... · ");
            StringAssert.StartsWith(MapSurface.Describe(player, 4500), "Teammate · ");
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-streamer.png"), PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class OneServer : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([new("Rustoria Main", "1.2.3.4", 28082) { ProfileKey = "key" }]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
