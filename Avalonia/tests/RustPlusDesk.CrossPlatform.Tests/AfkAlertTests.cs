using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class AfkAlertTests
{
    private static readonly DateTime T0 = new(2026, 10, 18, 12, 0, 0, DateTimeKind.Utc);

    private static ServerTeamMember M(string id, string name, double? x, double? y, bool online = true, bool alive = true) =>
        new(id, name, online, alive) { X = x, Y = y };

    [TestMethod]
    public void Tracker_ShowsAfkAfterFiveMinutes_AlertsOnceAtTheThreshold_AndReportsTheReturn()
    {
        var tracker = new AfkTracker();
        tracker.Observe([M("1", "Ann", 100, 100)], T0);
        Assert.AreEqual(0, tracker.Tick(T0.AddSeconds(299), 10).Count);
        Assert.IsNull(tracker.AfkText("1"));

        Assert.AreEqual(0, tracker.Tick(T0.AddMinutes(5), 10).Count, "The badge shows at 5 minutes, the alert waits for the threshold.");
        Assert.AreEqual("AFK: 5:00", tracker.AfkText("1"));
        tracker.Observe([M("1", "Ann", 100.04, 100)], T0.AddMinutes(6));
        Assert.AreEqual("AFK: 6:05", (tracker.Tick(T0.AddSeconds(365), 10), tracker.AfkText("1")).Item2, "Moves of 0.05 m or less are no movement.");

        var alert = tracker.Tick(T0.AddMinutes(10), 10).Single();
        Assert.IsFalse(alert.Returned);
        Assert.AreEqual("Ann", alert.Member.Name);
        Assert.AreEqual(0, tracker.Tick(T0.AddMinutes(11), 10).Count, "One alert per idle stretch.");

        tracker.Observe([M("1", "Ann", 140, 100)], T0.AddMinutes(72).AddSeconds(30));
        var back = tracker.Tick(T0.AddMinutes(72).AddSeconds(31), 10).Single();
        Assert.IsTrue(back.Returned);
        Assert.AreEqual(TimeSpan.FromMinutes(72.5), back.Duration);
        Assert.IsNull(tracker.AfkText("1"), "Moving clears the AFK text.");
        Assert.AreEqual(0, tracker.Tick(T0.AddMinutes(73), 10).Count, "The return is reported once.");
    }

    [TestMethod]
    public void Tracker_OfflineOrDeadMembersAreNeverAfk_AndMissingMembersLeaveAfterThreeReads()
    {
        var tracker = new AfkTracker();
        tracker.Observe([M("1", "Off", 1, 1, online: false), M("2", "Dead", 1, 1, alive: false), M("3", "Ann", null, null), M("0", "Zero", 1, 1)], T0);
        Assert.AreEqual("Ann", tracker.Tick(T0.AddMinutes(30), 5).Single().Member.Name, "Only the online, alive member goes AFK.");
        Assert.IsNull(tracker.AfkText("1"));
        Assert.IsNull(tracker.AfkText("2"));
        Assert.IsNull(tracker.AfkText("0"), "Steam ID 0 is skipped as in the reference.");
        Assert.AreEqual("AFK: 30:00", tracker.AfkText("3"), "Without a position the idle clock keeps running.");

        tracker.Observe([M("1", "Off", 1, 1)], T0.AddMinutes(31));
        Assert.AreEqual(0, tracker.Tick(T0.AddMinutes(32), 5).Count, "Coming online restarts the idle clock.");
        tracker.Observe([M("1", "Off", 1, 1)], T0.AddMinutes(33));
        Assert.AreEqual("AFK: 32:00", tracker.AfkText("3"), "Two missed reads keep the member.");
        tracker.Observe([M("1", "Off", 1, 1)], T0.AddMinutes(35));
        Assert.IsNull(tracker.AfkText("3"), "The third missed read drops the member.");

        tracker.Reset();
        Assert.AreEqual(0, tracker.Tick(T0.AddHours(2), 5).Count);
        Assert.AreEqual(5, AfkTracker.NormalizeThreshold(0));
        Assert.AreEqual(12, AfkTracker.NormalizeThreshold(12));
    }

    [TestMethod]
    public void Messages_UseTheReferenceTemplates()
    {
        var member = M("1", "Ann", 1, 1);
        Assert.AreEqual("Ann is AFK: 15:00", AfkTracker.Message(new(member, false, TimeSpan.FromMinutes(15)), "Ann", 15));
        Assert.AreEqual("Ann is not AFK anymore. Was AFK for 01:05", AfkTracker.Message(new(member, true, TimeSpan.FromMinutes(65.9)), "Ann", 15));
    }

    [TestMethod]
    public void ChatAlerts_AfkFlagsAndThresholdUseTheReferenceKeys()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.AfkAlertMinutes, 20);
        var model = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), new MapViewModel(), null, null));
        Assert.AreEqual(20, model.AfkAlertMinutes);
        CollectionAssert.AreEqual(new[] { "5 min", "10 min", "15 min", "20 min", "30 min" }, model.AfkThresholds.Select(item => item.Header).ToArray());
        CollectionAssert.AreEqual(new[] { false, false, false, true, false }, model.AfkThresholds.Select(item => item.IsSelected).ToArray());
        model.AfkThresholds[1].SelectCommand.Execute(null);
        Assert.AreEqual(10, settings.Values[SettingKeys.AfkAlertMinutes]);
        Assert.IsTrue(model.AfkThresholds[1].IsSelected && !model.AfkThresholds[3].IsSelected);
        Assert.IsFalse(model.IsMasterOn, "Choosing a threshold does not touch the flags.");

        model.ToggleMasterCommand.Execute(null);
        model.TogglePlayerAfkReturnCommand.Execute(null);
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnouncePlayerAfkReturn]);
        Assert.IsTrue(model.AnyPlayerAfk && !model.AnnouncePlayerAfk);
        Assert.AreEqual(1, model.EnabledCount);
        model.TogglePlayerAfkReturnCommand.Execute(null);
        Assert.IsFalse(model.IsMasterOn, "The last flag off turns the master off.");

        var empty = new ShopNoticeTests.MemorySettings();
        empty.Set(SettingKeys.AfkAlertMinutes, -3);
        Assert.AreEqual(5, new ChatAlertsViewModel(empty, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), new MapViewModel(), null, null)).AfkAlertMinutes);
    }

    [TestMethod]
    public async Task Team_TickUpdatesRows_AndTheAnnouncerPostsWithStreamerNamesOnlyForEnabledFlags() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), map, null, null));
        DateTime now = T0;
        var team = new TeamViewModel { Now = () => now };
        var sent = new List<string>();
        _ = new PresenceAnnouncer(team, map, alerts, () => null) { SendTeamChat = text => { sent.Add(text); return Task.CompletedTask; } };
        IReadOnlyList<ServerTeamMember> roster = [M("1", "Ann", 10, 10), M("2", "Bob", 20, 20)];
        using (var stop = new CancellationTokenSource())
        {
            Task running = team.RunAsync(new TeamViewModelTests.Session { Read = _ => Task.FromResult(roster) }, stop.Token);
            stop.Cancel();
            await running;
        }
        Assert.AreEqual(2, team.Members.Count);

        alerts.AfkThresholds[1].SelectCommand.Execute(null);
        now = T0.AddMinutes(5);
        team.TickAfk();
        Assert.AreEqual(0, sent.Count, "Alerts are off by default.");
        Assert.IsTrue(team.Members.All(row => row.IsAfk && row.AfkText == "AFK: 5:00"));

        alerts.ToggleMasterCommand.Execute(null);
        alerts.TogglePlayerAfkCommand.Execute(null);
        alerts.TogglePlayerAfkReturnCommand.Execute(null);
        map.AbbreviateNames = true;
        now = T0.AddMinutes(10);
        team.TickAfk();
        await Task.Yield();
        CollectionAssert.AreEqual(new[] { "A... is AFK: 10:00", "B... is AFK: 10:00" }, sent.ToArray());

        sent.Clear();
        alerts.TogglePlayerAfkCommand.Execute(null);
        roster = [M("1", "Ann", 50, 10), M("2", "Bob", 20, 20)];
        using (var stop = new CancellationTokenSource())
        {
            Task running = team.RunAsync(new TeamViewModelTests.Session { Read = _ => Task.FromResult(roster) }, stop.Token);
            stop.Cancel();
            await running;
        }
        now = T0.AddMinutes(10).AddSeconds(1);
        team.TickAfk();
        await Task.Yield();
        CollectionAssert.AreEqual(new[] { "A... is not AFK anymore. Was AFK for 00:10" }, sent.ToArray());
        Assert.IsFalse(team.Members[0].IsAfk);
        Assert.IsTrue(team.Members[1].IsAfk);

        team.Reset();
        Assert.AreEqual(0, team.Members.Count);
        team.TickAfk();
        Assert.AreEqual(1, sent.Count, "A reset forgets every idle clock.");
        return true;
    }, default);

    [TestMethod]
    public async Task TeamDock_ShowsTheAfkTextNextToTheState() => await UiTestSession.Current.Dispatch(async () =>
    {
        var model = new TeamViewModel
        {
            IsExpanded = true,
            Message = "Current server team",
            Members = [new TeamMemberRow(M("1", "Synthetic Builder", 1, 1), AfkText: "AFK: 12:34"),
                new TeamMemberRow(M("2", new string('W', 80), 1, 1), AfkText: "AFK: 125:07"),
                new TeamMemberRow(M("3", "Offline Farmer", 1, 1, online: false))]
        };
        var window = new Window { Width = 360, Height = 260, RequestedThemeVariant = ThemeVariant.Dark, Content = new TeamView { DataContext = model } };
        try
        {
            window.Show();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-team-afk.png"), PngBitmapEncoderOptions.Default);
            }
            var texts = window.GetVisualDescendants().OfType<TextBlock>().ToList();
            var afk = texts.Where(text => text.Text?.StartsWith("AFK:", StringComparison.Ordinal) == true).ToList();
            Assert.AreEqual(2, afk.Count(text => text.IsEffectivelyVisible));
            foreach (var text in afk)
            {
                var row = (Grid)text.Parent!;
                Assert.IsTrue(text.Bounds.Right <= row.Bounds.Width + 0.5, "The AFK text stays inside its row.");
                var name = row.Children.OfType<TextBlock>().First();
                Assert.IsTrue(name.Bounds.Right <= text.Bounds.Left, "The name trims before the AFK text.");
            }
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);
}
