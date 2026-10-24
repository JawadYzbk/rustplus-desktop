using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ChatViewModelTests
{
    internal static ChatMessage Message(int number, ChatChannel channel = ChatChannel.Team)
        => new(channel, 42, "Fixture player", $"Message {number}", DateTimeOffset.FromUnixTimeSeconds(1700000000 + number));

    [TestMethod]
    public void History_IsDeduplicatedSortedSearchableAndPaged_AndChannelsStaySeparate()
    {
        var model = new ServerChatViewModel();
        for (int i = 39; i >= 0; i--) { model.Team.Receive(Message(i), true); model.Team.Receive(Message(i), true); }
        model.Team.Receive(Message(100, ChatChannel.Clan));
        Assert.AreEqual(20, model.Team.Messages.Count);
        Assert.AreEqual("Message 20", model.Team.Messages[0].Text);
        Assert.IsTrue(model.Team.HasMore);
        model.Team.LoadMoreCommand.Execute(null);
        Assert.AreEqual(40, model.Team.Messages.Count);
        Assert.IsFalse(model.Team.HasMore);
        model.Team.Search = "message 3";
        Assert.AreEqual(11, model.Team.Messages.Count);
        model.Team.Search = "FIXTURE";
        Assert.AreEqual(20, model.Team.Messages.Count);
        Assert.AreEqual(0, model.Clan.Messages.Count);
        Assert.AreEqual(0, model.Team.UnreadCount);
    }

    [TestMethod]
    public void Unread_IsChannelSpecificAndBounded_AndClearedWhenReadOrProfileChanges()
    {
        var model = new ServerChatViewModel();
        for (int i = 0; i < 1100; i++) model.Clan.Receive(Message(i, ChatChannel.Clan));
        Assert.AreEqual(1000, model.Clan.UnreadCount);
        Assert.AreEqual("Clan (1000)", model.Clan.Title);
        model.SelectedIndex = 1;
        Assert.AreEqual(0, model.Clan.UnreadCount);
        model.Clan.Search = "Message 0";
        Assert.AreEqual(0, model.Clan.Messages.Count);
        model.IsVisible = false;
        model.Clan.Receive(Message(1101, ChatChannel.Clan));
        Assert.AreEqual(1, model.Clan.UnreadCount);
        model.IsVisible = true;
        Assert.AreEqual(0, model.Clan.UnreadCount);
        model.Team.Draft = "private draft";
        model.Clear();
        Assert.AreEqual("", model.Team.Draft);
        Assert.AreEqual(0, model.Clan.Messages.Count);
    }

    [TestMethod]
    public async Task Send_UsesSelectedChannelAndServerConfirmation_DoesNotEraseNewDraft()
    {
        var client = new TestChatClient();
        var result = new TaskCompletionSource<ChatMessage?>(TaskCreationOptions.RunContinuationsAsynchronously);
        client.Send = (channel, text, _) =>
        {
            Assert.AreEqual(ChatChannel.Clan, channel);
            Assert.AreEqual("hello", text);
            return result.Task;
        };
        var model = new ChatLaneViewModel(ChatChannel.Clan);
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        try
        {
            model.Draft = " hello ";
            var sending = model.SendCommand.ExecuteAsync(null);
            Assert.IsTrue(model.IsSending);
            Assert.IsFalse(model.SendCommand.CanExecute(null));
            model.Draft = "next message";
            var message = Message(1, ChatChannel.Clan);
            result.SetResult(message);
            await sending;
            Assert.AreEqual("next message", model.Draft);
            Assert.AreEqual(message, model.Messages.Single());
            Assert.AreEqual("Message sent.", model.Status);
            model.Receive(message);
            Assert.AreEqual(1, model.Messages.Count);
        }
        finally { stop.Cancel(); await running; }
        Assert.IsFalse(model.SendCommand.CanExecute(null));
    }

    [TestMethod]
    public async Task FailedSend_PreservesDraftAndDoesNotInventMessagesOrAutomaticallyRetry()
    {
        int sends = 0;
        var client = new TestChatClient { Send = (_, _, _) => { sends++; throw new Exception("private credential"); } };
        var model = new ChatLaneViewModel(ChatChannel.Team);
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        try
        {
            model.Draft = "keep me";
            await model.SendCommand.ExecuteAsync(null);
            Assert.AreEqual("keep me", model.Draft);
            Assert.AreEqual(0, model.Messages.Count);
            Assert.AreEqual(1, sends);
            Assert.IsFalse(model.Status.Contains("private"));
            client.Send = (_, _, _) => Task.FromResult<ChatMessage?>(null);
            await model.SendCommand.ExecuteAsync(null);
            Assert.AreEqual("", model.Draft);
            Assert.AreEqual(0, model.Messages.Count);
            StringAssert.Contains(model.Status, "Waiting for the server echo");
        }
        finally { stop.Cancel(); await running; }
    }

    [TestMethod]
    public async Task ProfileChange_CancelsPendingReadsAndSend_AndDiscardsLateOldCompletions()
    {
        var pendingRead = new TaskCompletionSource<IReadOnlyList<ChatMessage>>(TaskCreationOptions.RunContinuationsAsynchronously);
        var pendingSend = new TaskCompletionSource<ChatMessage?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var client = new TestChatClient { Read = (_, _) => pendingRead.Task, Send = (_, _, _) => pendingSend.Task };
        var model = new ServerChatViewModel();
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        model.Team.Draft = "old server";
        var sending = model.Team.SendCommand.ExecuteAsync(null);
        stop.Cancel();
        model.Clear();
        await running.WaitAsync(TimeSpan.FromSeconds(2));
        await sending;
        pendingRead.SetResult([Message(1)]);
        pendingSend.SetResult(Message(2));
        Assert.AreEqual(0, model.Team.Messages.Count);
        Assert.AreEqual("", model.Team.Draft);
        Assert.IsFalse(client.Watching);
        Assert.IsFalse(model.Team.IsConnected);
    }

    [TestMethod]
    public async Task ClanHistoryFailure_DoesNotStopTeamOrLiveBroadcasts_RefreshRecovers()
    {
        var client = new TestChatClient { Read = (channel, _) => channel == ChatChannel.Team
            ? Task.FromResult<IReadOnlyList<ChatMessage>>([Message(1)]) : throw new Exception("secret server response") };
        var model = new ServerChatViewModel();
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        try
        {
            Assert.AreEqual(1, model.Team.Messages.Count);
            Assert.IsTrue(model.Team.IsConnected);
            StringAssert.Contains(model.Clan.Status, "membership");
            Assert.IsFalse(model.Clan.Status.Contains("secret"));
            client.Emit(Message(2));
            await Eventually(() => model.Team.Messages.Count == 2);
            client.Read = (channel, _) => Task.FromResult<IReadOnlyList<ChatMessage>>([Message(3, channel)]);
            await model.Clan.RefreshCommand.ExecuteAsync(null);
            Assert.AreEqual(1, model.Clan.Messages.Count);
            Assert.AreEqual(1, model.Clan.UnreadCount);
        }
        finally { stop.Cancel(); await running; }
    }

    [TestMethod]
    public async Task MessageValidation_RejectsBlankOversizedAndControlCharactersBeforeSending()
    {
        var model = new ChatLaneViewModel(ChatChannel.Team);
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(new TestChatClient { Send = (_, _, _) => throw new AssertFailedException() }, stop.Token);
        try
        {
            foreach (var text in new[] { " ", new string('a', 129), "line\nbreak", "tab\ttext" })
            {
                model.Draft = text;
                Assert.IsFalse(model.SendCommand.CanExecute(null));
                Assert.ThrowsException<ChatOperationException>(() => ChatInput.Validate(text));
            }
            model.Draft = new string('a', 128);
            Assert.IsTrue(model.SendCommand.CanExecute(null));
            Assert.AreEqual("128/128", model.CharacterCount);
        }
        finally { stop.Cancel(); await running; }
    }

    internal static async Task Eventually(Func<bool> ready)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (!ready()) await Task.Delay(10, timeout.Token);
    }
}
