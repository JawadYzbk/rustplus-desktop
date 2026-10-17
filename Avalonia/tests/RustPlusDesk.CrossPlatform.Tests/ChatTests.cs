using System.Runtime.CompilerServices;
using System.Threading.Channels;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ChatTests
{
    private static readonly DateTimeOffset T0 = new(2026, 10, 18, 12, 0, 0, TimeSpan.Zero);

    [TestMethod]
    public async Task Chat_LoadsHistory_DeduplicatesLiveMessages_AndDetectsClan()
    {
        var session = new ChatSession { HasClan = false };
        session.History[ChatChannel.Team] = [new(ChatChannel.Team, 1, "Ann", "hello", T0), new(ChatChannel.Team, 2, "Bob", "raid?", T0.AddSeconds(5))];
        var model = new ChatViewModel();
        using var cts = new CancellationTokenSource();
        var running = model.RunAsync(session, cts.Token);
        await CompanionListenerViewModelTests.Eventually(() => model.Lines.Count == 2);
        Assert.IsTrue(model.IsConnected);
        Assert.IsFalse(model.HasClan, "A server without clans keeps the clan tab hidden.");
        // The same message arriving live after history is not shown twice; a new one is appended.
        session.Live.Writer.TryWrite(new(ChatChannel.Team, 2, "Bob", "raid?", T0.AddSeconds(5.5)));
        session.Live.Writer.TryWrite(new(ChatChannel.Team, 3, "Cid", "omw", T0.AddSeconds(9)));
        await CompanionListenerViewModelTests.Eventually(() => model.Lines.Count == 3);
        Assert.AreEqual("Cid", model.Lines[^1].Name);
        session.Live.Writer.TryWrite(new(ChatChannel.Clan, 4, "Dee", "clan hi", T0.AddSeconds(10)));
        await CompanionListenerViewModelTests.Eventually(() => model.HasClan);
        model.IsClan = true;
        Assert.AreEqual("clan hi", model.Lines.Single().Text);
        cts.Cancel();
        await running.WaitAsync(TimeSpan.FromSeconds(4));
        Assert.IsFalse(model.IsConnected);
        model.Reset();
        Assert.AreEqual(0, model.Lines.Count);
    }

    [TestMethod]
    public async Task Chat_SendsOnTheSelectedChannel_EnforcesTheLimit_AndKeepsTheDraftOnFailure()
    {
        var session = new ChatSession();
        var model = new ChatViewModel();
        using var cts = new CancellationTokenSource();
        var running = model.RunAsync(session, cts.Token);
        await CompanionListenerViewModelTests.Eventually(() => model.IsConnected);
        Assert.IsFalse(model.SendCommand.CanExecute(null), "Empty drafts cannot be sent.");
        model.Draft = new string('x', 200);
        Assert.AreEqual(128, model.Draft.Length);
        Assert.AreEqual("128/128", model.CharacterCount);
        Assert.IsTrue(model.IsAtLimit);
        model.Draft = "  on my way  ";
        await model.SendCommand.ExecuteAsync(null);
        Assert.AreEqual((ChatChannel.Team, "on my way"), session.Sent.Single());
        Assert.AreEqual("", model.Draft);
        session.FailSend = true;
        model.IsClan = true;
        model.Draft = "help";
        await model.SendCommand.ExecuteAsync(null);
        Assert.AreEqual("help", model.Draft);
        StringAssert.StartsWith(model.Error, "Message not sent");
        session.FailSend = false;
        await model.SendCommand.ExecuteAsync(null);
        Assert.AreEqual((ChatChannel.Clan, "help"), session.Sent[^1]);
        Assert.AreEqual("", model.Error);
        cts.Cancel();
        await running.WaitAsync(TimeSpan.FromSeconds(4));
    }

    [TestMethod]
    public async Task ChatView_RendersAndSendsWithEnter()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var session = new ChatSession();
            session.History[ChatChannel.Team] = [new(ChatChannel.Team, 1, "Ann", "hello team", T0)];
            var model = new ChatViewModel();
            using var cts = new CancellationTokenSource();
            var running = model.RunAsync(session, cts.Token);
            var window = new Window { Width = 700, Height = 260, Content = new ChatView { DataContext = model } };
            try
            {
                window.Show();
                await CompanionListenerViewModelTests.Eventually(() => model.Lines.Count == 1);
                var input = window.GetVisualDescendants().OfType<TextBox>().Single();
                input.Text = "gg";
                input.Focus();
                window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, null);
                await CompanionListenerViewModelTests.Eventually(() => session.Sent.Count == 1);
                using var frame = window.CaptureRenderedFrame();
                Assert.IsNotNull(frame);
            }
            finally { cts.Cancel(); await running; window.Close(); }
            return true;
        }, default);
    }

    private sealed class ChatSession : IServerChatSession
    {
        public Dictionary<ChatChannel, IReadOnlyList<ChatMessage>> History { get; } = new() { [ChatChannel.Team] = [], [ChatChannel.Clan] = [] };
        public Channel<ChatMessage> Live { get; } = Channel.CreateUnbounded<ChatMessage>();
        public List<(ChatChannel, string)> Sent { get; } = [];
        public bool HasClan { get; init; } = true;
        public bool FailSend { get; set; }
        public IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Chat", 0, 0, 0);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public Task<IReadOnlyList<ChatMessage>> ReadChatAsync(ChatChannel channel, CancellationToken ct)
            => channel == ChatChannel.Clan && !HasClan ? Task.FromException<IReadOnlyList<ChatMessage>>(new ServerConnectionException(ConnectionFailure.Unreachable))
                : Task.FromResult(History[channel]);
        public Task SendChatAsync(ChatChannel channel, string text, CancellationToken ct)
        {
            if (FailSend) return Task.FromException(new ServerConnectionException(ConnectionFailure.Unreachable));
            Sent.Add((channel, text));
            return Task.CompletedTask;
        }
        public async IAsyncEnumerable<ChatMessage> WatchChatAsync([EnumeratorCancellation] CancellationToken ct)
        {
            await foreach (var message in Live.Reader.ReadAllAsync(ct)) yield return message;
        }
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
