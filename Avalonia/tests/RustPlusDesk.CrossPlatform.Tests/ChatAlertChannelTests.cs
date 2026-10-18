using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ChatAlertChannelTests
{
    [TestMethod]
    public async Task FlagStore_ReadsAndWritesOnlyThatProfilesField()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyProfileFlagStore();
        Assert.IsFalse(await store.GetAsync(fixture.Path, server.ProfileKey, ProfileFlags.ChatAlertsUseClanChannel, false, default));
        Assert.IsTrue(await store.GetAsync(fixture.Path, server.ProfileKey, ProfileFlags.ChatAlertsUseClanChannel, true, default), "Missing field: the fallback.");
        var before = await fixture.JsonAsync();
        await store.SetAsync(fixture.Path, server.ProfileKey, ProfileFlags.ChatAlertsUseClanChannel, true, default);
        Assert.IsTrue(await store.GetAsync(fixture.Path, server.ProfileKey, ProfileFlags.ChatAlertsUseClanChannel, false, default));
        var after = await fixture.JsonAsync();
        Assert.AreEqual(true, after[0]!["ChatAlertsUseClanChannel"]!.GetValue<bool>());
        after[0]!.AsObject().Remove("ChatAlertsUseClanChannel");
        Assert.IsTrue(System.Text.Json.Nodes.JsonNode.DeepEquals(before, after));
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => store.SetAsync(fixture.Path, "missing", ProfileFlags.ChatAlertsUseClanChannel, true, default));
    }

    [TestMethod]
    public async Task Alerts_FollowTheProfilesChannel_AndAutomatedMessagesUseIt() => await UiTestSession.Current.Dispatch(async () =>
    {
        var flags = new MemoryFlags();
        var log = new List<string>();
        string? key = "a";
        var alerts = new ChatAlertsViewModel(new ShopNoticeTests.MemorySettings(), new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), new MapViewModel(), null, null))
        {
            ProfileFlags = flags,
            Profile = () => key is null ? null : ("profiles.json", key),
            Log = log.Add
        };
        Assert.IsTrue(alerts.HasChannelChoice);
        await alerts.LoadChannelAsync();
        Assert.AreEqual((ChatChannel.Team, true), (alerts.AlertChannel, alerts.UseTeamChannel));

        await alerts.UseClanCommand.ExecuteAsync(null);
        Assert.AreEqual(true, flags.Values["a"]);
        Assert.AreEqual(ChatChannel.Clan, alerts.AlertChannel);

        key = "b";
        await alerts.LoadChannelAsync();
        Assert.AreEqual(ChatChannel.Team, alerts.AlertChannel, "Each profile has its own channel.");
        key = "a";
        await alerts.LoadChannelAsync();
        Assert.AreEqual(ChatChannel.Clan, alerts.AlertChannel);

        flags.Fail = true;
        await alerts.UseTeamCommand.ExecuteAsync(null);
        Assert.AreEqual(ChatChannel.Clan, alerts.AlertChannel, "A failed save keeps the previous choice.");
        Assert.AreEqual("[chat] The alert channel could not be saved to the server profile.", log.Single());

        var session = new ChatTests.ChatSession();
        var chat = new ChatViewModel { Alerts = alerts };
        using var stop = new CancellationTokenSource();
        var running = chat.RunAsync(session, stop.Token);
        for (int i = 0; i < 50 && !chat.IsConnected; i++) await Task.Delay(10);
        Assert.IsTrue(await chat.SendAutomatedAsync("Synthetic alert"));
        Assert.AreEqual((ChatChannel.Clan, "Synthetic alert"), session.Sent.Single());
        stop.Cancel();
        try { await running; } catch (OperationCanceledException) { }
        return true;
    }, default);

    private sealed class MemoryFlags : IProfileFlagStore
    {
        public Dictionary<string, bool> Values { get; } = [];
        public bool Fail { get; set; }
        public Task<bool> GetAsync(string profilesPath, string profileKey, string name, bool fallback, CancellationToken ct) =>
            Task.FromResult(Values.TryGetValue(profileKey, out bool value) ? value : fallback);
        public Task SetAsync(string profilesPath, string profileKey, string name, bool value, CancellationToken ct)
        {
            if (Fail) throw new ProfileChangedException();
            Values[profileKey] = value;
            return Task.CompletedTask;
        }
    }
}
