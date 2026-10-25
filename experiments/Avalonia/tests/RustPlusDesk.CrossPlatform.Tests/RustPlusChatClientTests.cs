using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusApi.Data;
using RustPlusApi.Data.Clans;
using RustPlusApi.Data.Events;
using RustPlusApi.Interfaces;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Infrastructure.Connections;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class RustPlusChatClientTests
{
    [TestMethod]
    public async Task History_MapsTypedTeamAndClanResponsesAndUtcTimes()
    {
        var (api, proxy) = ApiProxy.Create();
        using var client = new RustPlusChatClient(api, new RustRequestGate());
        var team = await client.ReadHistoryAsync(ChatChannel.Team, CancellationToken.None);
        Assert.AreEqual("team", team.Single().Text);
        Assert.AreEqual(ChatChannel.Team, team.Single().Channel);
        Assert.AreEqual(TimeSpan.Zero, team.Single().SentAt.Offset);
        var clan = await client.ReadHistoryAsync(ChatChannel.Clan, CancellationToken.None);
        Assert.AreEqual("clan", clan.Single().Text);
        Assert.AreEqual(ChatChannel.Clan, clan.Single().Channel);
        proxy.FailRead = true;
        var error = await Assert.ThrowsExceptionAsync<ChatOperationException>(() => client.ReadHistoryAsync(ChatChannel.Team, CancellationToken.None));
        Assert.AreEqual(ChatFailure.Unavailable, error.Failure);
        Assert.IsFalse(error.ToString().Contains("secret"));
    }

    [TestMethod]
    public async Task Send_ChecksAcknowledgementAndChannel_NoOptimisticClanMessage()
    {
        var (api, proxy) = ApiProxy.Create();
        using var client = new RustPlusChatClient(api, new RustRequestGate());
        var confirmed = await client.SendAsync(ChatChannel.Team, " hello ", CancellationToken.None);
        Assert.AreEqual("hello", confirmed!.Text);
        Assert.AreEqual(1, proxy.TeamSends);
        Assert.AreEqual(0, proxy.ClanSends);
        Assert.IsNull(await client.SendAsync(ChatChannel.Clan, "hello", CancellationToken.None));
        Assert.AreEqual(1, proxy.ClanSends);
        proxy.Reject = true;
        var error = await Assert.ThrowsExceptionAsync<ChatOperationException>(() => client.SendAsync(ChatChannel.Team, "hello", CancellationToken.None));
        Assert.AreEqual(ChatFailure.Rejected, error.Failure);
    }

    [TestMethod]
    public async Task InvalidAndCancelledSends_NeverReachSdk()
    {
        var (api, proxy) = ApiProxy.Create();
        using var client = new RustPlusChatClient(api, new RustRequestGate());
        foreach (var text in new[] { "", "  ", new string('a', 129), "one\ntwo" })
            await Assert.ThrowsExceptionAsync<ChatOperationException>(() => client.SendAsync(ChatChannel.Team, text, CancellationToken.None));
        await Assert.ThrowsExceptionAsync<ChatOperationException>(() => client.SendAsync((ChatChannel)99, "hello", CancellationToken.None));
        using var stop = new CancellationTokenSource();
        stop.Cancel();
        await Assert.ThrowsExceptionAsync<TaskCanceledException>(() => client.SendAsync(ChatChannel.Team, "hello", stop.Token));
        Assert.AreEqual(0, proxy.TeamSends + proxy.ClanSends);
    }

    [TestMethod]
    public async Task Broadcasts_AreBoundedAndUnsubscribedOnDispose()
    {
        var (api, proxy) = ApiProxy.Create();
        var client = new RustPlusChatClient(api, new RustRequestGate());
        for (int i = 0; i < 300; i++) proxy.EmitTeam($"message {i}");
        proxy.EmitClan("clan event");
        client.Dispose();
        var messages = new List<ChatMessage>();
        await foreach (var message in client.WatchAsync(CancellationToken.None)) messages.Add(message);
        Assert.AreEqual(256, messages.Count);
        Assert.AreEqual("clan event", messages[^1].Text);
        Assert.AreEqual(ChatChannel.Clan, messages[^1].Channel);
        Assert.IsNull(proxy.TeamReceived);
        Assert.IsNull(proxy.ClanReceived);
        Assert.IsNull(RustPlusChatClient.Map(ChatChannel.Team, 0, "", " ", DateTime.UtcNow));
    }

    public class ApiProxy : DispatchProxy
    {
        public bool FailRead { get; set; }
        public bool Reject { get; set; }
        public int TeamSends { get; private set; }
        public int ClanSends { get; private set; }
        public EventHandler<TeamMessageEventArg>? TeamReceived { get; private set; }
        public EventHandler<ClanMessageEventArg>? ClanReceived { get; private set; }
        private static readonly DateTime Time = new(2026, 10, 24, 12, 0, 0, DateTimeKind.Unspecified);
        public void EmitTeam(string text) => TeamReceived?.Invoke(this, new() { SteamId = 42, Name = "Fixture", Message = text, Time = Time });
        public void EmitClan(string text) => ClanReceived?.Invoke(this, new() { SteamId = 42, Name = "Fixture", Message = text, Time = Time });
        public static (IRustPlus Api, ApiProxy Proxy) Create()
        {
            var api = DispatchProxy.Create<IRustPlus, ApiProxy>();
            return (api, (ApiProxy)api);
        }
        protected override object? Invoke(MethodInfo? method, object?[]? args)
        {
            switch (method?.Name)
            {
                case "add_OnTeamChatReceived": TeamReceived += (EventHandler<TeamMessageEventArg>)args![0]!; return null;
                case "remove_OnTeamChatReceived": TeamReceived -= (EventHandler<TeamMessageEventArg>)args![0]!; return null;
                case "add_OnClanChatReceived": ClanReceived += (EventHandler<ClanMessageEventArg>)args![0]!; return null;
                case "remove_OnClanChatReceived": ClanReceived -= (EventHandler<ClanMessageEventArg>)args![0]!; return null;
                case "GetTeamChatAsync":
                    return Task.FromResult(new Response<TeamChatInfo?> { IsSuccess = !FailRead, Data = new()
                        { Messages = [new() { SteamId = 42, Name = "Fixture", Message = "team", Time = Time }] }, Error = new() { Message = "secret" } });
                case "GetClanChatAsync":
                    return Task.FromResult(new Response<ClanChatInfo?> { IsSuccess = true, Data = new()
                        { Messages = [new() { SteamId = 42, Name = "Fixture", Message = "clan", Time = Time }] } });
                case "SendTeamMessageAsync":
                    TeamSends++;
                    return Task.FromResult(new Response<TeamMessage?> { IsSuccess = !Reject,
                        Data = new() { SteamId = 42, Name = "Fixture", Message = (string)args![0]!, Time = Time } });
                case "SendClanMessageAsync": ClanSends++; return Task.FromResult(new Response { IsSuccess = !Reject });
                default: throw new AssertFailedException($"Unexpected SDK call: {method?.Name}");
            }
        }
    }
}
