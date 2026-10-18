using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Notifications;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Notifications;
using RustPlusDesk.Core.Settings;
using RustPlusDesk.Infrastructure.Notifications;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class NotificationTests
{
    private static readonly DateTime Now = new(2026, 10, 18, 12, 0, 0, DateTimeKind.Local);

    private static InboxNotification Row(string type, string title, string message = "m", int secondsAgo = 0, string? ip = "10.0.0.1", int? port = 28082) =>
        new() { Type = type, Title = title, Message = message, Timestamp = Now.AddSeconds(-secondsAgo), ServerIp = ip, ServerPort = port, ServerName = "Test Server" };

    private sealed class MemoryHistory : INotificationHistoryStore
    {
        public List<InboxNotification> Saved { get; set; } = [];
        public Task<IReadOnlyList<InboxNotification>> LoadAsync(CancellationToken ct) => Task.FromResult<IReadOnlyList<InboxNotification>>(Saved.ToList());
        public Task SaveAsync(IReadOnlyList<InboxNotification> items, CancellationToken ct) { Saved = items.ToList(); return Task.CompletedTask; }
    }

    [TestMethod]
    public void Inbox_AppliesReferenceDeduplicationRetentionMutesAndCap()
    {
        var inbox = new NotificationInbox();
        var muted = new List<string>();
        var first = Row("Alarm", "Raid", "Base alarm");
        first.ImageUrl = "https://example.invalid/logo.png";
        Assert.IsTrue(inbox.Add(first, Now, 30, muted));
        Assert.IsFalse(inbox.Add(Row("Alarm", "Raid", "Base alarm", secondsAgo: -3), Now, 30, muted), "Same alarm within four seconds.");
        var later = Row("Alarm", "Raid", "Base alarm", secondsAgo: -5);
        Assert.IsTrue(inbox.Add(later, Now, 30, muted));
        Assert.AreEqual("https://example.invalid/logo.png", later.ImageUrl, "Rows borrow the server's picture.");
        Assert.AreSame(later, inbox.Items[0], "Newest first.");
        var push = Row("Death", "You were killed by X");
        push.FcmNotificationId = "fcm-1";
        Assert.IsTrue(inbox.Add(push, Now, 30, muted));
        var again = Row("Death", "You were killed by X", "other", secondsAgo: -60);
        again.FcmNotificationId = "fcm-1";
        Assert.IsFalse(inbox.Add(again, Now, 30, muted), "The same push is kept once.");
        muted.Add("10.0.0.1:28082");
        Assert.IsFalse(inbox.Add(Row("Alarm", "Muted", "x", secondsAgo: -100), Now, 30, muted));
        Assert.IsTrue(inbox.Add(Row("Alarm", "No server", "x", ip: null, port: null), Now, 30, muted));
        Assert.AreEqual(4, inbox.UnreadCount);
        Assert.IsTrue(inbox.Add(Row("Alarm", "Old", "y", secondsAgo: 31 * 86400), Now, 30, []) is true);
        Assert.IsFalse(inbox.Items.Any(item => item.Title == "Old"), "Rows past the retention are dropped.");

        var big = new NotificationInbox();
        for (int i = 0; i < 505; i++) big.Add(Row("Alarm", "A", $"m{i}", secondsAgo: -i), Now, 30, []);
        Assert.AreEqual(NotificationInbox.MaxItems, big.Items.Count);
        Assert.AreEqual("m504", big.Items[0].Message);
        Assert.AreEqual(30, NotificationInbox.Retention(0));
        Assert.AreEqual(7, NotificationInbox.Retention(7));
    }

    [TestMethod]
    public void Inbox_LoadsRewordsFiltersAndEdits()
    {
        var oldDeath = Row("Death", "Offline Death", "x", secondsAgo: 60);
        oldDeath.AttackerName = "Bob";
        var oldPair = Row("Pairing", "Pairing Successful", "y", secondsAgo: 30);
        oldPair.EntityId = 77;
        oldPair.EntityName = "Front door";
        var expired = Row("Alarm", "Expired", secondsAgo: 9 * 86400);
        var inbox = new NotificationInbox();
        inbox.Load([oldDeath, expired, oldPair], Now, 7);
        Assert.AreEqual(2, inbox.Items.Count);
        Assert.AreEqual("Front door", inbox.Items[0].Title);
        Assert.AreEqual("Device paired (ID 77)", inbox.Items[0].Message);
        Assert.AreEqual("You were killed by Bob", inbox.Items[1].Title);
        Assert.AreEqual("Test Server", inbox.Items[1].Message);

        Assert.AreEqual(1, inbox.Filter("", NotificationFilter.Deaths).Count());
        Assert.AreEqual(0, inbox.Filter("", NotificationFilter.Alarms).Count());
        Assert.AreEqual(2, inbox.Filter(" test server ", NotificationFilter.All).Count(), "Server names match.");
        Assert.AreEqual(1, inbox.Filter("door", NotificationFilter.All).Count());
        Assert.IsTrue(inbox.MarkRead(inbox.Items[0].Id));
        Assert.IsFalse(inbox.MarkRead(inbox.Items[0].Id));
        Assert.IsTrue(inbox.MarkAllRead());
        Assert.IsFalse(inbox.MarkAllRead());
        Assert.IsTrue(inbox.Delete(inbox.Items[0].Id));
        inbox.Clear();
        Assert.AreEqual(0, inbox.Items.Count);
    }

    [TestMethod]
    public void FromPush_UsesTheReferenceWording()
    {
        var at = new DateTimeOffset(2026, 10, 18, 9, 0, 0, TimeSpan.Zero);
        var alarm = NotificationInbox.FromPush(new(CompanionUpdateKind.Alarm, "HV Raid wake up", "Your base is under attack!") { Host = "1.2.3.4", Port = 28082, ServerName = "Rusty", OccurredAt = at })!;
        Assert.AreEqual("Alarm", alarm.Type);
        Assert.AreEqual("HV Raid wake up", alarm.Title);
        Assert.AreEqual("1.2.3.4:28082", alarm.ServerKey);
        Assert.AreEqual(at.LocalDateTime, alarm.Timestamp);
        Assert.AreEqual("Alarm", NotificationInbox.FromPush(new(CompanionUpdateKind.Alarm, " ", "x"))!.Title);
        var death = NotificationInbox.FromPush(new(CompanionUpdateKind.Death, "You were killed by Bob", "Bob") { ServerName = "Rusty" })!;
        Assert.AreEqual("You were killed by Bob", death.Title);
        Assert.AreEqual("Rusty", death.Message);
        Assert.AreEqual("Bob", death.AttackerName);
        var device = NotificationInbox.FromPush(new(CompanionUpdateKind.Pairing, "Smart Switch", "Tap to pair") { EntityId = 5 })!;
        Assert.AreEqual("Smart Switch", device.Title);
        Assert.AreEqual("Device paired (ID 5)", device.Message);
        var server = NotificationInbox.FromPush(new(CompanionUpdateKind.Pairing, "Rusty", "Tap to pair") { Host = "1.2.3.4", Port = 28082 })!;
        Assert.AreEqual("1.2.3.4:28082", server.Title);
        Assert.AreEqual("Server paired.", server.Message);
        Assert.IsNull(NotificationInbox.FromPush(new(CompanionUpdateKind.Chat, "Team", "hi")));
        Assert.IsNull(NotificationInbox.FromPush(new(CompanionUpdateKind.Listening, "Listening", "")));
    }

    [TestMethod]
    public async Task Store_ReadsTheReferenceHistoryShape_AndBacksUpUnreadableFiles()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-notifications-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(folder, "cache", "notifications_history.json");
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.WriteAllTextAsync(path, """
                [ { "Id": "a1", "Timestamp": "2026-10-17T20:15:00+02:00", "Type": "Alarm", "Title": "Raid", "Message": "Base", "ServerIp": "1.2.3.4",
                    "ServerPort": 28082, "ServerName": "Rusty", "IsRead": false, "EntityId": 12, "EntityName": null, "AttackerName": null,
                    "ChatAuthor": null, "FcmNotificationId": "f1", "ImageUrl": null } ]
                """);
            var store = new NotificationHistoryStore(path);
            var item = (await store.LoadAsync(CancellationToken.None)).Single();
            Assert.AreEqual("a1", item.Id);
            Assert.AreEqual(12u, item.EntityId);
            Assert.AreEqual("1.2.3.4:28082", item.ServerKey);
            item.IsRead = true;
            await store.SaveAsync([item], CancellationToken.None);
            string json = await File.ReadAllTextAsync(path);
            StringAssert.Contains(json, "\"IsRead\": true");
            StringAssert.Contains(json, "\"FcmNotificationId\": \"f1\"");

            await File.WriteAllTextAsync(path, "[ broken");
            var broken = new NotificationHistoryStore(path);
            Assert.AreEqual(0, (await broken.LoadAsync(CancellationToken.None)).Count);
            await broken.SaveAsync([], CancellationToken.None);
            Assert.AreEqual(1, Directory.GetFiles(Path.GetDirectoryName(path)!, "notifications_history.json.unreadable-*.bak").Length);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public async Task ViewModel_TracksUnread_FiltersConfirmsAndMutes()
    {
        var history = new MemoryHistory { Saved = [Row("Alarm", "Saved alarm", secondsAgo: 120)] };
        var settings = new ShopNoticeTests.MemorySettings();
        var model = new NotificationsViewModel(history, settings, () => Now);
        await model.LoadAsync(CancellationToken.None);
        Assert.AreEqual(1, model.UnreadCount);
        Assert.IsTrue(model.HasUnread);
        Assert.IsTrue(model.Handle(new(CompanionUpdateKind.Death, "You were killed by Ann", "Ann") { Host = "10.0.0.2", Port = 28015, ServerName = "Other", OccurredAt = Now }));
        Assert.IsFalse(model.Handle(new(CompanionUpdateKind.Chat, "Team", "hi")), "Chat pushes are not inbox rows.");
        Assert.AreEqual(2, model.UnreadCount);
        Assert.AreEqual("You were killed by Ann", model.Rows[0].Title);
        Assert.IsTrue(model.Rows[0].IsDeath);
        await model.FlushAsync();
        Assert.AreEqual(2, history.Saved.Count);

        model.SetFilterCommand.Execute(NotificationFilter.Deaths);
        Assert.IsTrue(model.ShowDeaths);
        Assert.AreEqual(NotificationFilter.Deaths, model.Filter);
        Assert.AreEqual(1, model.Rows.Count);
        model.SetFilterCommand.Execute(NotificationFilter.All);
        model.SearchText = "saved";
        Assert.AreEqual("Saved alarm", model.Rows.Single().Title);
        model.SearchText = "";

        model.OpenCommand.Execute(model.Rows[0]);
        Assert.AreEqual(1, model.UnreadCount);
        model.MarkAllReadCommand.Execute(null);
        Assert.AreEqual(0, model.UnreadCount);
        Assert.IsFalse(model.Rows.Any(row => row.IsUnread));

        model.RequestMuteCommand.Execute(model.Rows[0]);
        Assert.IsTrue(model.IsMuteConfirmOpen);
        StringAssert.Contains(model.MuteQuestion, "Other");
        model.ConfirmMuteCommand.Execute(null);
        CollectionAssert.AreEqual(new[] { "10.0.0.2:28015" }, ((List<string>)settings.Values[SettingKeys.MutedNotificationServers]!).ToArray());
        Assert.AreEqual("Other", ((Dictionary<string, string>)settings.Values[SettingKeys.MutedNotificationServerNames]!)["10.0.0.2:28015"]);
        Assert.AreEqual("Notifications from Other are now muted.", model.Notice);
        Assert.IsFalse(model.Handle(new(CompanionUpdateKind.Alarm, "Muted", "x") { Host = "10.0.0.2", Port = 28015, OccurredAt = Now.AddMinutes(1) }));

        var targets = new List<string>();
        model.ConnectToServer = (host, port) => { targets.Add($"{host}:{port}"); return host == "10.0.0.1"; };
        model.ConnectServerCommand.Execute(model.Rows.Single(row => row.Title == "Saved alarm"));
        Assert.AreEqual("", model.Notice);
        model.ConnectServerCommand.Execute(model.Rows.Single(row => row.IsDeath));
        Assert.AreEqual("Could not find a paired server profile for 10.0.0.2:28015. Please pair this server first.", model.Notice);
        CollectionAssert.AreEqual(new[] { "10.0.0.1:28082", "10.0.0.2:28015" }, targets);

        model.DeleteCommand.Execute(model.Rows[0]);
        Assert.AreEqual(1, model.Rows.Count);
        model.RequestClearCommand.Execute(null);
        Assert.IsTrue(model.IsConfirmingClear);
        model.CancelClearCommand.Execute(null);
        Assert.AreEqual(1, model.Rows.Count);
        model.RequestClearCommand.Execute(null);
        model.ConfirmClearCommand.Execute(null);
        Assert.IsFalse(model.HasRows);
        await model.FlushAsync();
        Assert.AreEqual(0, history.Saved.Count);
        model.RequestClearCommand.Execute(null);
        Assert.IsFalse(model.IsConfirmingClear, "Nothing to clear.");
    }

    [TestMethod]
    public async Task NotificationsView_RendersRowsAndEmptyState()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new NotificationsViewModel(null, null, () => Now);
            var window = new Window { Width = 1000, Height = 700, Content = new Border { Padding = new global::Avalonia.Thickness(16), Child = new NotificationsView { DataContext = model } } };
            try
            {
                window.Show();
                Assert.IsTrue(window.GetVisualDescendants().OfType<TextBlock>().Any(text => text.Text == "No notifications here" && text.IsEffectivelyVisible));
                model.Handle(new(CompanionUpdateKind.Alarm, "HV Rockets Raid wake up", "Your base is under attack!") { Host = "10.0.0.1", Port = 28082, ServerName = "Test Server", OccurredAt = Now });
                model.Handle(new(CompanionUpdateKind.Death, "You were killed by Ann", "Ann") { Host = "10.0.0.1", Port = 28082, ServerName = "Test Server", OccurredAt = Now.AddMinutes(-3) });
                model.Handle(new(CompanionUpdateKind.Pairing, "Smart Switch", "Tap to pair") { EntityId = 42, Host = "10.0.0.1", Port = 28082, ServerName = "Test Server", OccurredAt = Now.AddMinutes(-9) });
                model.OpenCommand.Execute(model.Rows.Single(row => row.Title == "Smart Switch"));
                window.UpdateLayout();
                var rows = window.GetVisualDescendants().OfType<Button>().Where(button => button.Classes.Contains("notification")).ToList();
                Assert.AreEqual(3, rows.Count);
                CollectionAssert.AreEqual(new[] { "HV Rockets Raid wake up", "You were killed by Ann", "Smart Switch" }, model.Rows.Select(row => row.Title).ToArray(),
                    "Newest first by time, whatever the arrival order.");
                Assert.AreEqual(2, rows.Count(row => row.Classes.Contains("unread")));
                var mute = rows[0].GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Mute notification server");
                mute.Command!.Execute(mute.CommandParameter);
                window.UpdateLayout();
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-notifications.png"), PngBitmapEncoderOptions.Default);
                }
                Assert.IsTrue(model.IsMuteConfirmOpen);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }
}
