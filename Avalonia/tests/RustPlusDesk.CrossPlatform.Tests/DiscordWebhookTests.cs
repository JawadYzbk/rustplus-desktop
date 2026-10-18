using System.Net;
using System.Text.Json;
using System.Text.Json.Nodes;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Alerts;
using RustPlusDesk.Infrastructure.Cloud;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DiscordWebhookTests
{
    private const string Url = "https://discord.example.test/api/webhooks/1/synthetic";

    [TestMethod]
    public void Settings_FollowTheReferenceSetters_AndContent()
    {
        var settings = new DiscordWebhookSettings().WithUrl(Url);
        Assert.IsTrue(settings.Enabled, "A URL turns the webhook on.");
        Assert.IsFalse(settings.WithUrl("  ").Enabled, "An empty URL turns it off.");
        settings = settings.WithMention("@here", true).WithMention("@everyone", true);
        Assert.AreEqual("@here @everyone", settings.Mention);
        Assert.IsTrue(settings.MentionsEveryone && settings.MentionsHere);
        Assert.AreEqual("@everyone", settings.WithMention("@here", false).Mention);
        Assert.AreSame(settings, settings.WithMention("@here", true));
        Assert.AreEqual("@here\n**[Rusty]** Heli spawned", DiscordWebhookSettings.Content("@here", "Rusty", "Heli spawned"));
        Assert.AreEqual("**[Rust Server]** Hi", DiscordWebhookSettings.Content(" ", null, "Hi"));
    }

    [TestMethod]
    public async Task Store_WritesTheFiveReferenceFieldsOnly()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDiscordWebhookStore();
        Assert.AreEqual(new DiscordWebhookSettings(), await store.GetAsync(fixture.Path, server.ProfileKey, default));
        var before = await fixture.JsonAsync();
        var settings = new DiscordWebhookSettings(Url, "@here", true, true, true);
        await store.SetAsync(fixture.Path, server.ProfileKey, settings, default);
        Assert.AreEqual(settings, await store.GetAsync(fixture.Path, server.ProfileKey, default));
        var after = (await fixture.JsonAsync())[0]!.AsObject();
        foreach (string field in new[] { "Url", "Mention", "Enabled", "Tts", "Exclusive" })
            Assert.IsTrue(after.Remove("DiscordWebhookChatAlerts" + field), field);
        Assert.IsTrue(JsonNode.DeepEquals(before[0], after));
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => store.SetAsync(fixture.Path, "missing", settings, default));
    }

    [TestMethod]
    public async Task Sender_PostsContentAndTts_AndSignInReadsTheStoredSession()
    {
        var bodies = new List<string>();
        var status = HttpStatusCode.NoContent;
        using var http = new HttpClient(new Handler(async request => { bodies.Add(await request.Content!.ReadAsStringAsync()); return new HttpResponseMessage(status); }));
        var sender = new HttpDiscordWebhookSender(http);
        await sender.SendAsync(Url, "**[Rusty]** Hi", true, default);
        Assert.AreEqual("{\"content\":\"**[Rusty]** Hi\",\"tts\":true}", bodies.Single());
        status = HttpStatusCode.NotFound;
        await Assert.ThrowsExceptionAsync<HttpRequestException>(() => sender.SendAsync(Url, "x", false, default));

        string folder = Path.Combine(Path.GetTempPath(), "rpd-signin-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            string path = Path.Combine(folder, "cloud_desktop_token.json");
            var state = new CloudSignInState(path);
            Assert.IsFalse(await state.IsSignedInAsync(default));
            File.WriteAllText(path, JsonSerializer.Serialize(new { Token = "synthetic-cloud-token", ExpiresAt = "2050-01-01T00:00:00Z" }));
            Assert.IsTrue(await state.IsSignedInAsync(default));
            File.WriteAllText(path, JsonSerializer.Serialize(new { Token = "synthetic-cloud-token", ExpiresAt = "2000-01-01T00:00:00Z" }));
            Assert.IsFalse(await state.IsSignedInAsync(default));
            File.WriteAllText(path, "{ broken");
            Assert.IsFalse(await state.IsSignedInAsync(default));
        }
        finally { Directory.Delete(folder, true); }
    }

    [TestMethod]
    public async Task Alerts_CountTheWebhook_NeedSignIn_AndDiscordOnlySkipsInGameChat() => await UiTestSession.Current.Dispatch(async () =>
    {
        var store = new MemoryHooks();
        var sender = new RecordingSender();
        var signIn = new FakeSignIn();
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), map, null, null));
        var log = new List<string>();
        var hook = new DiscordWebhookViewModel(store, sender, signIn, () => ("profiles.json", "key", "Rusty")) { Log = log.Add };
        alerts.Webhook = hook;
        await hook.LoadAsync();
        Assert.IsFalse(hook.IsSignedIn);
        Assert.IsTrue(hook.SignInNote.Length > 0);

        hook.UrlDraft = "not a url";
        await hook.SaveUrlCommand.ExecuteAsync(null);
        Assert.IsFalse(store.Values.ContainsKey("key"));
        hook.UrlDraft = Url;
        await hook.SaveUrlCommand.ExecuteAsync(null);
        Assert.AreEqual((Url, true), (store.Values["key"].Url, store.Values["key"].Enabled));
        Assert.AreEqual(1, alerts.EnabledCount, "The webhook counts like an alert flag.");
        Assert.IsFalse(alerts.CanToggleWebhook, "Master off and signed out.");

        await hook.SendAsync("Synthetic alert");
        Assert.AreEqual(0, sender.Sent.Count, "Signed out: nothing is posted.");
        await hook.ToggleEnabledCommand.ExecuteAsync(null);
        Assert.IsFalse(hook.Enabled);
        await hook.ToggleEnabledCommand.ExecuteAsync(null);
        Assert.IsFalse(hook.Enabled, "Switching on needs a cloud sign-in.");

        signIn.SignedIn = true;
        await hook.LoadAsync();
        await hook.ToggleEnabledCommand.ExecuteAsync(null);
        Assert.IsTrue(hook.Enabled);
        alerts.ToggleMasterCommand.Execute(null);
        Assert.IsTrue(alerts.CanToggleWebhook && alerts.AnnounceWebhook);
        hook.MentionEveryone = true;
        hook.Tts = true;
        await Task.Yield();
        Assert.AreEqual("@everyone", store.Values["key"].Mention);

        var session = new ChatTests.ChatSession();
        var chat = new ChatViewModel { Alerts = alerts };
        using var stop = new CancellationTokenSource();
        var running = chat.RunAsync(session, stop.Token);
        for (int i = 0; i < 50 && !chat.IsConnected; i++) await Task.Delay(10);
        await chat.SendAutomatedAsync("Heli spawned", "🚁 Heli spawned");
        await Task.Yield();
        Assert.AreEqual((Url, "@everyone\n**[Rusty]** 🚁 Heli spawned", true), sender.Sent.Single());
        Assert.AreEqual((ChatChannel.Team, "Heli spawned"), session.Sent.Single());

        hook.Exclusive = true;
        Assert.IsFalse(alerts.ChannelChoiceEnabled, "Discord only disables the Team/Clan choice.");
        await chat.SendAutomatedAsync("Chinook spawned");
        await Task.Yield();
        Assert.AreEqual(2, sender.Sent.Count);
        Assert.AreEqual(1, session.Sent.Count, "Discord only: not posted in game.");

        sender.Fail = true;
        await hook.SendAsync("x");
        Assert.AreEqual("[Discord] Webhook send failed.", log.Single());

        alerts.DeselectAllCommand.Execute(null);
        await Task.Yield();
        Assert.IsFalse(hook.Enabled);
        Assert.IsFalse(alerts.IsMasterOn, "Nothing left on, the webhook included.");
        alerts.SelectAllCommand.Execute(null);
        await Task.Yield();
        Assert.IsTrue(hook.Enabled);

        await hook.ClearUrlCommand.ExecuteAsync(null);
        Assert.AreEqual(("", false), (store.Values["key"].Url, store.Values["key"].Enabled));
        stop.Cancel();
        try { await running; } catch (OperationCanceledException) { }
        return true;
    }, default);

    [TestMethod]
    public async Task AfkAlerts_UseTheReferenceDiscordTexts() => await UiTestSession.Current.Dispatch(async () =>
    {
        var map = new MapViewModel();
        var alerts = new ChatAlertsViewModel(new ShopNoticeTests.MemorySettings(), new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), map, null, null));
        var team = new TeamViewModel();
        var sent = new List<(string, string)>();
        var announcer = new PresenceAnnouncer(team, map, alerts, () => null) { SendTeamChatWithDiscordText = (text, discord) => { sent.Add((text, discord)); return Task.CompletedTask; } };
        alerts.ToggleMasterCommand.Execute(null);
        alerts.TogglePlayerAfkCommand.Execute(null);
        alerts.TogglePlayerAfkReturnCommand.Execute(null);
        var member = new ServerTeamMember("1", "Ann", true, true);
        announcer.ObserveAfk(new(member, false, TimeSpan.FromMinutes(5)));
        announcer.ObserveAfk(new(member, true, TimeSpan.FromMinutes(70)));
        await Task.Yield();
        CollectionAssert.AreEqual(new[] { ("Ann is AFK: 5:00", "💤 Ann is AFK: 5:00"), ("Ann is not AFK anymore. Was AFK for 01:10", "🏃 Ann is not AFK anymore. Was AFK for 01:10") },
            sent.ToArray());
        return true;
    }, default);

    [TestMethod]
    public async Task SettingsCard_Renders() => await UiTestSession.Current.Dispatch(async () =>
    {
        var hook = new DiscordWebhookViewModel(new MemoryHooks(), new RecordingSender(), new FakeSignIn(), () => ("profiles.json", "key", "Rusty"));
        await hook.LoadAsync();
        var model = new SettingsViewModel(new ShopNoticeTests.MemorySettings()) { Webhook = hook };
        var window = new Window { Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark,
            Content = new Border { Padding = new Thickness(16), Child = new SettingsView { DataContext = model } } };
        try
        {
            window.Show();
            model.FocusWebhook();
            window.UpdateLayout();
            var box = window.GetVisualDescendants().OfType<TextBox>().Single(t => AutomationProperties.GetName(t) == "Discord webhook URL");
            Assert.IsTrue(box.IsEffectivelyVisible);
            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-settings-webhook.png"), PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    private sealed class MemoryHooks : IDiscordWebhookStore
    {
        public Dictionary<string, DiscordWebhookSettings> Values { get; } = [];
        public Task<DiscordWebhookSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct) => Task.FromResult(Values.GetValueOrDefault(profileKey) ?? new());
        public Task SetAsync(string profilesPath, string profileKey, DiscordWebhookSettings settings, CancellationToken ct) { Values[profileKey] = settings; return Task.CompletedTask; }
    }

    private sealed class RecordingSender : IDiscordWebhookSender
    {
        public List<(string, string, bool)> Sent { get; } = [];
        public bool Fail { get; set; }
        public Task SendAsync(string url, string content, bool tts, CancellationToken ct)
        {
            if (Fail) throw new HttpRequestException("synthetic");
            Sent.Add((url, content, tts));
            return Task.CompletedTask;
        }
    }

    private sealed class FakeSignIn : ICloudSignInState
    {
        public bool SignedIn { get; set; }
        public Task<bool> IsSignedInAsync(CancellationToken ct) => Task.FromResult(SignedIn);
    }

    private sealed class Handler(Func<HttpRequestMessage, Task<HttpResponseMessage>> send) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => send(request);
    }
}
