using System.Net;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Infrastructure.Ai;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference the AI companion: the chat command that asks it, the rules that decide who may, and the provider call.
/// </summary>
[TestClass]
public sealed class AiCompanionTests
{
    [TestMethod]
    public void TheCommandMatchesOnTheFirstWord_AndTakesTheQuestion()
    {
        Assert.IsTrue(ChatAiCommand.Matches("ai what is a rifle", "ai"));
        Assert.IsTrue(ChatAiCommand.Matches("ai", "ai"));
        Assert.IsFalse(ChatAiCommand.Matches("air", "ai"), "A command that only starts with the word is not it.");
        Assert.IsFalse(ChatAiCommand.Matches("pop", "ai"));
        Assert.IsFalse(ChatAiCommand.Matches("ai", ""), "Without a command name there is no AI command.");

        Assert.AreEqual("what is a rifle", ChatAiCommand.Question("!ai what is a rifle", "!", "ai"));
        Assert.AreEqual("What is a rifle?", ChatAiCommand.Question("!ai What is a rifle?", "!", "ai"));
        Assert.AreEqual("", ChatAiCommand.Question("!ai", "!", "ai"));
    }

    [TestMethod]
    public void TheAnswerIsCutAtASentenceEnd()
    {
        string answer = "A rifle is a weapon. It fires bullets at the enemies. " + new string('x', 300);
        string shortened = ChatAiCommand.Shorten(answer);
        Assert.IsTrue(shortened.Length <= ChatAiCommand.MaxChars + 1);
        Assert.IsTrue(shortened.EndsWith('.') || shortened.EndsWith('…'),
            "The answer is cut at a sentence end, or marked as cut.");
    }

    [TestMethod]
    public void TheAnswerIsCleanedOfTheModelsQuoting()
    {
        Assert.AreEqual("Hello.", ChatAiCommand.Clean("\"Hello.\""));
        Assert.AreEqual("line one", ChatAiCommand.Clean("```\nline one\n```"));
    }

    [TestMethod]
    public void TheAllowanceIsARollingHour()
    {
        DateTime now = new(2026, 10, 20, 12, 0, 0, DateTimeKind.Utc);
        var asks = new List<DateTime> { now.AddMinutes(-30), now.AddMinutes(-70) };
        Assert.IsTrue(ChatAiCommand.HasAllowanceLeft(5, asks, now), "Only the ask inside the hour counts.");
        Assert.IsFalse(ChatAiCommand.HasAllowanceLeft(1, asks, now));
        Assert.IsTrue(ChatAiCommand.HasAllowanceLeft(0, asks, now), "Zero means no limit.");
    }

    [TestMethod]
    public void TheCommandIsListedOnlyWhenItCanBeUsed()
    {
        Assert.IsFalse(ChatAiCommand.IsListable(isSetUp: false, isPremium: true, isOwner: true, allowTeammates: true),
            "Without a key there is nothing to ask.");
        Assert.IsFalse(ChatAiCommand.IsListable(isSetUp: true, isPremium: false, isOwner: true, allowTeammates: true),
            "Asking from chat is a supporter feature.");
        Assert.IsTrue(ChatAiCommand.IsListable(isSetUp: true, isPremium: true, isOwner: true, allowTeammates: false));
        Assert.IsFalse(ChatAiCommand.IsListable(isSetUp: true, isPremium: true, isOwner: false, allowTeammates: false));
        Assert.IsTrue(ChatAiCommand.IsListable(isSetUp: true, isPremium: true, isOwner: false, allowTeammates: true));
    }

    [TestMethod]
    public async Task TheProviderAsksAndReadsTheAnswer()
    {
        var handler = new StubHandler(HttpStatusCode.OK, """
            {"choices":[{"message":{"content":"A rifle is a weapon."}}]}
            """);
        using var http = new HttpClient(handler);
        var provider = new OpenAiProvider(http);
        var answer = await provider.AskAsync(new AiQuestion { Text = "what is a rifle" }, "secret-key", null, default);

        Assert.AreEqual("A rifle is a weapon.", answer.Answer);
        Assert.AreEqual("/v1/chat/completions", handler.LastRequest!.RequestUri!.AbsolutePath);
        Assert.AreEqual("Bearer secret-key", handler.LastRequest!.Headers.Authorization!.ToString());
    }

    [TestMethod]
    public async Task ARefusedRequest_SaysSo()
    {
        var handler = new StubHandler(HttpStatusCode.Unauthorized, """{"error":"bad key"}""");
        using var http = new HttpClient(handler);
        var provider = new OpenAiProvider(http);
        await Assert.ThrowsExceptionAsync<AiRequestException>(() =>
            provider.AskAsync(new AiQuestion { Text = "hello" }, "bad", null, default));
    }

    [TestMethod]
    public async Task TheSettingsKeepTheKeyInAPrivateFile()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-ai-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, "ai-companion.json");
        try
        {
            var store = new AiSettingsStore(path);
            Assert.IsFalse((await store.LoadAsync(default)).IsSetUp);

            await store.SaveAsync(new AiSettings { Provider = AiProviders.OpenRouter }
                .WithModel(AiProviders.OpenRouter, "gpt-4o")
                .WithKey(AiProviders.OpenRouter, "sk-test"), default);
            var loaded = await store.LoadAsync(default);
            Assert.IsTrue(loaded.IsSetUp);
            Assert.AreEqual("sk-test", loaded.ApiKey);
            Assert.AreEqual(AiProviders.OpenRouter, loaded.Provider);
            Assert.AreEqual("gpt-4o", loaded.Model);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public async Task TheChatCommandAsksTheCompanion_BeforeTheEngine()
    {
        var chat = new ChatViewModel();
        var asked = new List<string>();
        var service = new ChatCommandService(chat)
        {
            Settings = ChatCommandSettings.Defaults with { Ai = "ai" },
            AskAi = async (message, question, ct) =>
            {
                asked.Add(question);
                return "A rifle is a weapon.";
            }
        };

        await service.HandleAsync(new ChatMessage(ChatChannel.Team, 42, "Alice", "!ai what is a rifle", DateTimeOffset.UtcNow), default);
        Assert.AreEqual(1, asked.Count);
        Assert.AreEqual("what is a rifle", asked[0]);
    }

    [TestMethod]
    public async Task TheChatCommandAsksForAQuestion_WhenThereIsNone()
    {
        var chat = new ChatViewModel();
        var asked = new List<string>();
        var service = new ChatCommandService(chat)
        {
            Settings = ChatCommandSettings.Defaults with { Ai = "ai" },
            AskAi = async (message, question, ct) =>
            {
                asked.Add(question);
                return "A rifle is a weapon.";
            }
        };

        await service.HandleAsync(new ChatMessage(ChatChannel.Team, 42, "Alice", "!ai", DateTimeOffset.UtcNow), default);
        Assert.AreEqual(0, asked.Count, "An empty question is answered by the app, not by the companion.");
    }

    [TestMethod]
    public async Task AnotherCommand_IsNotSentToTheCompanion()
    {
        var chat = new ChatViewModel();
        var asked = new List<string>();
        var service = new ChatCommandService(chat)
        {
            Settings = ChatCommandSettings.Defaults with { Ai = "ai" },
            AskAi = async (message, question, ct) =>
            {
                asked.Add(question);
                return "A rifle is a weapon.";
            }
        };

        await service.HandleAsync(new ChatMessage(ChatChannel.Team, 42, "Alice", "!pop", DateTimeOffset.UtcNow), default);
        Assert.AreEqual(0, asked.Count);
    }

    private sealed class StubHandler(HttpStatusCode status, string body) : HttpMessageHandler
    {
        public HttpRequestMessage? LastRequest { get; private set; }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            LastRequest = request;
            return Task.FromResult(new HttpResponseMessage(status) { Content = new StringContent(body, Encoding.UTF8, "application/json") });
        }
    }
}
