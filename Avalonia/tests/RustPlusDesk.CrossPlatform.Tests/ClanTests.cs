using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Clans;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Clans;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference IsClanCommandAuthorAllowed (MainWindow.Map.ChatCommands) and the clan tab
/// (MainWindow.Clan.Core): a clan command is answered only for an author whose clan role the owner granted, the
/// role list follows the live clan, and the gate runs before the prefix and the cooldown.
/// </summary>
[TestClass]
public sealed class ClanTests
{
    private static readonly DateTime Now = new(2026, 5, 1, 12, 0, 0, DateTimeKind.Utc);
    private const ulong AuthorId = 76561198000000001;
    private const ulong StrangerId = 76561198000000002;
    private const ulong OutsiderId = 76561198000000003;

    private static ChatCommandReply Run(ChatChannel channel, string text, ChatCommandState state,
        ChatCommandSettings settings, ChatCommandCooldown? cooldown = null, ulong author = AuthorId) =>
        ChatCommands.Process(new ChatCommandRequest(channel, "Alice", author, text), state, settings,
            cooldown ?? new ChatCommandCooldown(), Now);

    private static ChatCommandSettings ClanEnabled => ChatCommandSettings.Defaults with
    {
        ClanCommandsEnabled = true,
        ClanRolePermissions = [new ChatRolePermission(7, "Owner", 100, true)]
    };

    private static ChatCommandState WithClan => new(Connected: true, Players: 4,
        ClanMembers: [new ChatClanMember(AuthorId, 7, "Owner"), new ChatClanMember(StrangerId, 9, "Member")]);

    // ---- the gate -------------------------------------------------------------------------------

    [TestMethod]
    public void ClanCommands_AreIgnored_UntilTheOwnerTurnsThemOn()
    {
        var settings = ChatCommandSettings.Defaults with { ClanCommandsEnabled = false };
        var reply = Run(ChatChannel.Clan, "!pop", WithClan, settings);
        Assert.IsNull(reply.Message);
        Assert.IsNull(reply.Log, "The reference refuses a clan command silently when clan answering is off.");
    }

    [TestMethod]
    public void AnAuthorOutsideTheClan_IsIgnoredWithTheReferenceLogLine()
    {
        var reply = Run(ChatChannel.Clan, "!pop", WithClan, ClanEnabled, author: OutsiderId);
        Assert.IsNull(reply.Message);
        Assert.AreEqual("[ChatCommand] Clan command from Alice ignored: author not in the last clan pull.", reply.Log);
    }

    [TestMethod]
    public void ARoleThatWasNeverGranted_IsIgnoredWithTheReferenceLogLine()
    {
        var settings = ClanEnabled with { ClanRolePermissions = [new ChatRolePermission(7, "Owner", 100, false)] };
        var reply = Run(ChatChannel.Clan, "!pop", WithClan, settings);
        Assert.IsNull(reply.Message);
        Assert.AreEqual("[ChatCommand] Clan command from Alice ignored: role 'Owner' is not permitted.", reply.Log);
    }

    [TestMethod]
    public void AGrantedRole_AnswersInClanChat()
    {
        var reply = Run(ChatChannel.Clan, "!pop", WithClan, ClanEnabled);
        Assert.AreEqual("4 players currently online.", reply.Message);
    }

    [TestMethod]
    public void TeamChat_IsNeverGatedByTheClanRules()
    {
        var settings = ChatCommandSettings.Defaults with { ClanCommandsEnabled = false };
        Assert.AreEqual("4 players currently online.", Run(ChatChannel.Team, "!pop", WithClan, settings).Message,
            "Being in the team is the permission.");
    }

    [TestMethod]
    public void TheClanGate_RunsBeforeThePrefixAndTheCooldown()
    {
        var cooldown = new ChatCommandCooldown();
        var refused = Run(ChatChannel.Clan, "pop", WithClan, ClanEnabled, cooldown, StrangerId);
        Assert.IsNull(refused.Message);
        Assert.IsNotNull(Run(ChatChannel.Team, "!pop", WithClan, ClanEnabled, cooldown).Message,
            "A refused clan message does not spend the cooldown.");
    }

    [TestMethod]
    public void WithoutAClanPull_EveryClanAuthorIsRefused()
    {
        var reply = Run(ChatChannel.Clan, "!pop", new ChatCommandState(Connected: true, Players: 4), ClanEnabled);
        Assert.IsNull(reply.Message);
        Assert.AreEqual("[ChatCommand] Clan command from Alice ignored: author not in the last clan pull.", reply.Log);
    }

    // ---- the settings card ----------------------------------------------------------------------

    [TestMethod]
    public async Task TheRoleRows_FollowTheLiveClan_AndKeepTheGrantedTicks()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new RustPlusDesk.Infrastructure.Profiles.LegacyChatCommandStore();
        var settings = ClanEnabled with
        {
            ClanRolePermissions = [new ChatRolePermission(7, "Owner", 100, true), new ChatRolePermission(9, "Member", 50, false)]
        };
        await store.SaveAsync(fixture.Path, server.ProfileKey, settings, default);

        var roles = new List<ClanRole> { new(9, "Member", 50), new(7, "Owner", 100), new(11, "Newcomer", 10) };
        var card = new ChatCommandSettingsViewModel(store, () => (fixture.Path, server), clanRoles: () => roles);
        await card.LoadAsync();

        CollectionAssert.AreEqual(new[] { "Newcomer", "Member", "Owner" }, card.ClanRoles.Select(row => row.Name).ToArray(),
            "The rows follow the clan's own order, lowest rank first.");
        Assert.IsTrue(card.ClanRoles.Single(row => row.RoleId == 7).Allowed, "The granted role keeps its tick.");
        Assert.IsFalse(card.ClanRoles.Single(row => row.RoleId == 9).Allowed);
        Assert.IsFalse(card.ClanRoles.Single(row => row.RoleId == 11).Allowed, "A new role arrives switched off.");
        Assert.IsTrue(card.ClanCommandsEnabled);
    }

    [TestMethod]
    public async Task TheClanFields_RoundTripThroughTheProfileStore()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new RustPlusDesk.Infrastructure.Profiles.LegacyChatCommandStore();
        var settings = ChatCommandSettings.Defaults with
        {
            ClanCommandsEnabled = true,
            ClanCommandsAllowAi = true,
            ClanCommandsAllowBaseCodes = true,
            ClanCommandsAllowSwitches = true,
            ClanRolePermissions = [new ChatRolePermission(7, "Owner", 100, true), new ChatRolePermission(9, "Member", 50, false)]
        };
        await store.SaveAsync(fixture.Path, server.ProfileKey, settings, default);

        var read = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual(settings.ClanCommandsEnabled, read.ClanCommandsEnabled);
        Assert.AreEqual(settings.ClanCommandsAllowAi, read.ClanCommandsAllowAi);
        Assert.AreEqual(settings.ClanCommandsAllowBaseCodes, read.ClanCommandsAllowBaseCodes);
        Assert.AreEqual(settings.ClanCommandsAllowSwitches, read.ClanCommandsAllowSwitches);
        CollectionAssert.AreEqual(settings.ClanRolePermissions.ToArray(), read.ClanRolePermissions.ToArray());
    }

    // ---- the clan tab ---------------------------------------------------------------------------

    [TestMethod]
    public void TheClanTab_ShowsTheClan_AndTellsTheEmptyAndFailedStatesApart()
    {
        var model = new ClanViewModel();
        Assert.IsFalse(model.HasClan);
        Assert.AreEqual("Connect to a server to see its clan.", model.Message);

        model.Apply(new ClanSnapshot("7", "Rustlings", "Welcome", [new ClanRole(7, "Owner", 100)],
            [new ClanMember(AuthorId, "Alice", 7, "Owner", true), new ClanMember(StrangerId, "Bob", 9, "Member", false)], 100), false);
        Assert.IsTrue(model.HasClan);
        Assert.AreEqual("Rustlings", model.Name);
        Assert.AreEqual("Welcome", model.Motd);
        Assert.AreEqual("1 online · 2 / 100", model.MemberCountText);
        Assert.AreEqual(2, model.Members.Count);
        Assert.AreEqual("Alice", model.Members[0].Name);

        model.Apply(null, false);
        Assert.IsFalse(model.HasClan);
        Assert.AreEqual("You are not in a clan.", model.Message);

        model.Apply(null, true);
        Assert.AreEqual("The clan could not be read. The list may be out of date.", model.Message);
    }

    // ---- the workspace ---------------------------------------------------------------------------

    [TestMethod]
    public async Task TheClanWorkspace_OpensFromTheRail_AndShowsTheMembers()
    {
        using var profiles = new DeviceProfileEditorTests.Fixture();
        await UiTestSession.Current.Dispatch(async () =>
        {
            var session = new ClanSession();
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), profiles.Path, new Factory(session));
            var window = new RustPlusDesk.Avalonia.Shell.MainWindow { DataContext = model, Width = 1000, Height = 700 };
            try
            {
                await model.RefreshCommand.ExecuteAsync(null);
                model.Connection.ConnectCommand.Execute(null);
                window.Show();

                model.OpenClanCommand.Execute(null);
                Assert.IsTrue(model.IsClanSelected);
                var view = window.GetVisualDescendants().OfType<RustPlusDesk.Avalonia.Features.Clans.ClanView>().Single();
                Assert.IsTrue(view.IsVisible);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-clan.png"), PngBitmapEncoderOptions.Default);
                }

                Assert.IsTrue(model.Connection.Clan.HasClan, "The clan is polled while the connection lives.");
                Assert.AreEqual("Rustlings", model.Connection.Clan.Name);
                Assert.AreEqual(2, model.Connection.Clan.Members.Count);
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, default);
    }

    private sealed class ClanSession : IServerTeamSession, IClanSession
    {
        public IReadOnlyList<ServerTeamMember> Members { get; } = [new("76561198000000001", "Self")];
        public IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Fixture", 1, 100, 0);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken ct) => Task.FromResult(InitialStatus);
        public Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct) => Task.FromResult(Members);
        public Task<ClanSnapshot?> ReadClanAsync(CancellationToken ct) => Task.FromResult<ClanSnapshot?>(new ClanSnapshot("7", "Rustlings", "Welcome",
            [new ClanRole(7, "Owner", 100), new ClanRole(9, "Member", 50)],
            [new ClanMember(76561198000000001, "Alice", 7, "Owner", true),
             new ClanMember(76561198000000002, "Bob", 9, "Member", false)], 100));
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }

    private sealed class Picker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    private sealed class Factory(IServerSession session) : RustPlusDesk.Core.Connections.IServerSessionFactory
    {
        public Task<IServerSession> ConnectAsync(string path, string key, CancellationToken ct) => Task.FromResult(session);
    }
}
