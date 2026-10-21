using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Infrastructure.Ai;
using RustPlusDesk.Infrastructure.Speech;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference AiCompanionSettings' voice: the answer read out loud, the language it is in, and the game data that goes
/// with the question.
/// </summary>
[TestClass]
public sealed class AiVoiceTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-ai-voice-" + Guid.NewGuid().ToString("N"));
        public AiSettingsStore Store { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            Store = new AiSettingsStore(Path.Combine(Folder, "ai-companion.json"));
        }

        public void Dispose()
        {
            try { if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true); }
            catch (IOException) { }
        }
    }

    [TestMethod]
    public async Task TheVoiceChoicesAreSavedWithTheKeys()
    {
        using var fixture = new Fixture();

        var settings = new AiSettings { Provider = AiProviders.OpenAi }
            .WithKey(AiProviders.OpenAi, "sk-test")
            .WithVoice(audioAnswers: false, textAnswers: true, includeGameData: false, answerLanguage: AiAnswerLanguages.English);

        await fixture.Store.SaveAsync(settings, CancellationToken.None);

        var loaded = await fixture.Store.LoadAsync(CancellationToken.None);
        Assert.IsFalse(loaded.AudioAnswers, "The answer is not read out loud.");
        Assert.IsTrue(loaded.TextAnswers, "The answer is still written down.");
        Assert.IsFalse(loaded.IncludeGameData, "The game data is not sent.");
        Assert.AreEqual(AiAnswerLanguages.English, loaded.AnswerLanguage, "The answer is in English.");
        Assert.AreEqual("sk-test", loaded.KeyFor(AiProviders.OpenAi), "The key is still there.");
    }

    [TestMethod]
    public async Task TheCardSavesTheVoiceChoices()
    {
        using var fixture = new Fixture();
        var card = new AiSettingsViewModel(fixture.Store);
        await card.LoadAsync();

        card.AudioAnswers = false;
        card.IncludeGameData = false;
        card.AnswerLanguage = AiAnswerLanguages.App;
        await card.SaveCommand.ExecuteAsync(null);

        var loaded = await fixture.Store.LoadAsync(CancellationToken.None);
        Assert.IsFalse(loaded.AudioAnswers, "The card's switch was written.");
        Assert.IsFalse(loaded.IncludeGameData, "The card's game-data switch was written.");
        Assert.AreEqual(AiAnswerLanguages.App, loaded.AnswerLanguage, "The card's answer language was written.");
    }

    [TestMethod]
    public void TheThreeAnswerLanguagesAreTheReferenceOnes()
    {
        Assert.AreEqual(3, AiAnswerLanguages.All.Length, "The reference offers three languages for an answer.");
        Assert.IsTrue(AiAnswerLanguages.All.Contains(AiAnswerLanguages.MatchQuestion), "The question's own language is one of them.");
        Assert.IsTrue(AiAnswerLanguages.All.Contains(AiAnswerLanguages.App), "The app's own language is one of them.");
        Assert.IsTrue(AiAnswerLanguages.All.Contains(AiAnswerLanguages.English), "English is one of them.");
    }

    [TestMethod]
    public void ThePlatformVoiceSaysSoWhenThereIsNone()
    {
        var speech = PlatformSpeech.ForCurrentPlatform();

        if (OperatingSystem.IsWindows())
        {
            // A machine with no voice installed is a machine with no voice, and the app says so.
            Assert.IsTrue(speech.UnavailableReason.Length > 0 || speech.IsAvailable,
                "The voice either works or says why it does not.");
        }
        else
        {
            Assert.IsFalse(speech.IsAvailable, "There is no voice to ask for on this platform.");
            Assert.IsTrue(speech.UnavailableReason.Length > 0, "The reason is given.");
        }
    }
}
