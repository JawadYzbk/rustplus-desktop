using System.Runtime.CompilerServices;
using System.Threading.Channels;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Controls.Primitives;
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
                var input = window.GetVisualDescendants().OfType<TextBox>().Single(box => AutomationProperties.GetName(box) == "Chat message");
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

    [TestMethod]
    public void Chat_SearchFiltersByTextOrAuthor_IgnoringCase_AndClosingClearsIt()
    {
        var model = new ChatViewModel();
        model.Add(new(ChatChannel.Team, 1, "Ann", "Raid at B4", T0));
        model.Add(new(ChatChannel.Team, 2, "Bob", "omw", T0.AddSeconds(1)));
        model.Add(new(ChatChannel.Clan, 3, "Cid", "raid clan", T0.AddSeconds(2)));
        model.ToggleSearchCommand.Execute(null);
        Assert.IsTrue(model.IsSearchOpen);
        model.SearchText = "  RAID ";
        Assert.AreEqual("Ann", model.Lines.Single().Name, "Only the current channel is searched.");
        model.SearchText = "bob";
        Assert.AreEqual("omw", model.Lines.Single().Text, "The author name matches too.");
        model.Add(new(ChatChannel.Team, 2, "Bob", "second", T0.AddSeconds(3)));
        Assert.AreEqual(2, model.Lines.Count, "New matching messages appear while searching.");
        model.SearchText = "nothing";
        Assert.IsTrue(model.HasNoMatches);
        model.ClearSearchCommand.Execute(null);
        Assert.AreEqual(3, model.Lines.Count);
        Assert.IsFalse(model.HasNoMatches);
        model.SearchText = "omw";
        model.ToggleSearchCommand.Execute(null);
        Assert.IsFalse(model.IsSearchOpen);
        Assert.AreEqual("", model.SearchText);
        Assert.AreEqual(3, model.Lines.Count);
        model.ToggleSearchCommand.Execute(null);
        model.SearchText = "omw";
        model.Reset();
        Assert.IsFalse(model.IsSearchOpen, "A new connection starts without a search.");
        Assert.AreEqual("", model.SearchText);
    }

    [TestMethod]
    public void Chat_MessageMenu_MentionsTheAuthor_AndCentresTheMapOnThem()
    {
        var centred = new List<ulong>();
        var model = new ChatViewModel { CenterOnMap = id => { centred.Add(id); return id == 7; } };
        model.Add(new(ChatChannel.Team, 7, "Ann", "here", T0));
        model.Add(new(ChatChannel.Team, 8, "Bob", "there", T0.AddSeconds(1)));
        model.Add(new(ChatChannel.Team, 0, "", "server", T0.AddSeconds(2)));
        var (ann, bob, system) = (model.Lines[0], model.Lines[1], model.Lines[2]);
        model.Draft = "hi";
        model.Mention(ann);
        Assert.AreEqual("@Ann hi", model.Draft);
        model.Mention(system);
        Assert.AreEqual("@Ann hi", model.Draft, "A line without an author name has nothing to mention.");
        model.Draft = new string('x', 127);
        model.Mention(bob);
        Assert.AreEqual(128, model.Draft.Length, "The in-game limit still applies.");
        Assert.AreEqual("https://steamcommunity.com/profiles/7", ann.SteamProfile!.ToString());
        Assert.IsFalse(system.HasSteamId);
        Assert.IsNull(system.SteamProfile);
        model.CenterOnAuthorCommand.Execute(bob);
        Assert.AreEqual("Teammate is not currently visible on the active map.", model.Notice);
        model.CenterOnAuthorCommand.Execute(ann);
        Assert.AreEqual("", model.Notice);
        model.CenterOnAuthorCommand.Execute(system);
        CollectionAssert.AreEqual(new ulong[] { 8, 7 }, centred);
    }

    [TestMethod]
    public async Task ChatView_SearchBoxFiltersAndTheMessageMenuMentions()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var session = new ChatSession();
            session.History[ChatChannel.Team] = [new(ChatChannel.Team, 1, "Ann", "hello team", T0), new(ChatChannel.Team, 2, "Bob", "raid now", T0.AddSeconds(1))];
            var model = new ChatViewModel();
            using var cts = new CancellationTokenSource();
            var running = model.RunAsync(session, cts.Token);
            var window = new Window { Width = 700, Height = 260, Content = new ChatView { DataContext = model } };
            try
            {
                window.Show();
                await CompanionListenerViewModelTests.Eventually(() => model.Lines.Count == 2);
                var search = window.GetVisualDescendants().OfType<TextBox>().Single(box => AutomationProperties.GetName(box) == "Search chat");
                Assert.IsFalse(search.IsEffectivelyVisible);
                var toggle = window.GetVisualDescendants().OfType<ToggleButton>().Single(button => AutomationProperties.GetName(button) == "Toggle chat search");
                toggle.Command!.Execute(null);
                window.UpdateLayout();
                Assert.IsTrue(search.IsEffectivelyVisible);
                search.Text = "raid";
                window.UpdateLayout();
                var line = window.GetVisualDescendants().OfType<TextBlock>().Single(block => block.DataContext is ChatLine && block.ContextMenu is not null);
                Assert.AreEqual("Bob", ((ChatLine)line.DataContext!).Name);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-chat-search.png"), global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                line.ContextMenu!.Open(line);
                var items = line.ContextMenu.Items.OfType<MenuItem>().ToList();
                CollectionAssert.AreEqual(new[] { "Copy text", "Copy Steam ID", "Mention @User", "Center on map", "Open Steam profile" },
                    items.Select(item => (string)item.Header!).ToArray());
                items.Single(item => (string)item.Header! == "Mention @User").RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                Assert.AreEqual("@Bob ", model.Draft);
                line.ContextMenu.Close();
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
