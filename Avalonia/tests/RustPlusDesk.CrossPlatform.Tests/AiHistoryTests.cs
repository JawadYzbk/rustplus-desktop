using System.Text.Json;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Ai;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Infrastructure.Ai;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference AiHistory and AiHistoryWindow: the last few questions asked of the AI companion, kept on this machine.
///
/// The file is the reference's own, so the WPF app and this one read each other's history, and the window is where a
/// failure that happened while the answer panel was switched off is still readable afterwards.
/// </summary>
[TestClass]
public sealed class AiHistoryTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(),
            "rustplus-ai-history-" + Guid.NewGuid().ToString("N"));
        public string FilePath { get; }

        public Fixture(string contents = "")
        {
            FilePath = Path.Combine(Folder, "ai-history.json");
            Directory.CreateDirectory(Folder);
            if (contents.Length > 0) File.WriteAllText(FilePath, contents);
        }

        public void Dispose()
        {
            if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
        }
    }

    private static AiExchange Exchange(string answer) => new()
    {
        WhenUtc = new DateTime(2026, 10, 20, 12, 0, 0, DateTimeKind.Utc),
        Provider = AiProviders.OpenAi,
        Model = "gpt-4o-mini",
        Transcript = "what is a rifle",
        Answer = answer
    };

    // ---- the file -----------------------------------------------------------------------------

    [TestMethod]
    public async Task TheExchangesRoundTripThroughTheFile()
    {
        using var file = new Fixture();
        var store = new LegacyAiHistoryStore(file.FilePath);
        store.Add(new AiExchange
        {
            WhenUtc = new DateTime(2026, 10, 20, 12, 0, 0, DateTimeKind.Utc),
            Provider = AiProviders.OpenRouter,
            Model = "gpt-4o",
            Transcript = "what is a rifle",
            HadScreenshot = true,
            HadGameAudio = true,
            Answer = "A rifle is a weapon."
        });

        var read = new LegacyAiHistoryStore(file.FilePath);
        Assert.AreEqual(1, read.Entries.Count, "The exchange is read back.");
        var entry = read.Entries[0];
        Assert.AreEqual("what is a rifle", entry.Transcript);
        Assert.AreEqual("A rifle is a weapon.", entry.Answer);
        Assert.AreEqual(AiProviders.OpenRouter, entry.Provider);
        Assert.AreEqual("gpt-4o", entry.Model);
        Assert.IsTrue(entry.HadScreenshot, "The screenshot is remembered.");
        Assert.IsTrue(entry.HadGameAudio, "The game's own sound is remembered.");
        Assert.AreEqual(new DateTime(2026, 10, 20, 12, 0, 0, DateTimeKind.Utc), entry.WhenUtc);
        Assert.IsFalse(entry.Failed, "An answered exchange did not fail.");

        string json = await File.ReadAllTextAsync(file.FilePath);
        Assert.IsTrue(json.Contains("\n"), "The reference's file is written indented.");
    }

    [TestMethod]
    public async Task TheFileKeepsTheReferencePropertyNames()
    {
        using var file = new Fixture();
        var store = new LegacyAiHistoryStore(file.FilePath);
        store.Add(Exchange("A rifle is a weapon."));

        string json = await File.ReadAllTextAsync(file.FilePath);
        // The eight the reference writes; Failed is worked out from the error when it is read back.
        foreach (string name in new[] { "WhenUtc", "Provider", "Model", "Transcript", "HadScreenshot", "HadGameAudio", "Answer", "Error" })
            Assert.IsTrue(json.Contains(name, StringComparison.Ordinal), $"The reference's file has {name} in it.");

        using var document = JsonDocument.Parse(json);
        Assert.AreEqual(JsonValueKind.Array, document.RootElement.ValueKind, "The reference's file is a list.");
    }

    [TestMethod]
    public void TheNewestIsFirst()
    {
        using var file = new Fixture();
        var store = new LegacyAiHistoryStore(file.FilePath);
        store.Add(Exchange("first"));
        store.Add(Exchange("second"));

        Assert.AreEqual("second", store.Entries[0].Answer, "The newest exchange is at the front.");
        Assert.AreEqual("first", store.Entries[1].Answer);
    }

    [TestMethod]
    public void OnlyTheLastTwentyAreKept()
    {
        using var file = new Fixture();
        var store = new LegacyAiHistoryStore(file.FilePath);
        for (int index = 0; index < AiHistory.MaxEntries + 5; index++) store.Add(Exchange($"answer {index}"));

        Assert.AreEqual(AiHistory.MaxEntries, store.Entries.Count, "Twenty is the limit the reference keeps.");
        Assert.AreEqual($"answer {AiHistory.MaxEntries + 4}", store.Entries[0].Answer, "The newest is kept.");
        Assert.AreEqual("answer 5", store.Entries[^1].Answer, "The oldest five are dropped.");
    }

    [TestMethod]
    public void AMissingFileGivesAnEmptyHistory()
    {
        using var file = new Fixture();
        var store = new LegacyAiHistoryStore(file.FilePath);

        Assert.AreEqual(0, store.Entries.Count, "Nothing asked yet is an empty history, not a broken app.");
        Assert.IsFalse(File.Exists(file.FilePath), "Nothing is written until there is something to write.");
    }

    [TestMethod]
    public void ABrokenFileGivesAnEmptyHistory()
    {
        using var file = new Fixture("{ this is not json at all");
        var store = new LegacyAiHistoryStore(file.FilePath);

        Assert.AreEqual(0, store.Entries.Count, "A file that cannot be read is an empty history, not a broken app.");
    }

    [TestMethod]
    public async Task ClearingEmptiesTheFile()
    {
        using var file = new Fixture();
        var store = new LegacyAiHistoryStore(file.FilePath);
        store.Add(Exchange("A rifle is a weapon."));

        store.Clear();

        Assert.AreEqual(0, store.Entries.Count, "The history is empty.");
        string json = await File.ReadAllTextAsync(file.FilePath);
        Assert.AreEqual("[]", json.Trim(), "The reference's file is emptied, not deleted.");
    }

    // ---- the app's own path -------------------------------------------------------------------

    [TestMethod]
    public void TheHistorySaysWhenItChanges()
    {
        using var file = new Fixture();
        AiHistory.Use(new LegacyAiHistoryStore(file.FilePath));
        int changes = 0;
        void Count(object? sender, EventArgs e) => changes++;
        AiHistory.Changed += Count;
        try
        {
            AiHistory.Add(Exchange("A rifle is a weapon."));
            Assert.AreEqual(1, changes, "An added exchange is announced.");
            AiHistory.Clear();
            Assert.AreEqual(2, changes, "A cleared history is announced.");
        }
        finally
        {
            AiHistory.Changed -= Count;
            AiHistory.Use(null);
        }
    }

    // ---- the window ----------------------------------------------------------------------------

    [TestMethod]
    public async Task TheWindowListsTheExchanges_AndShowsWhyOneFailed() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var file = new Fixture();
        AiHistory.Use(new LegacyAiHistoryStore(file.FilePath));
        AiHistory.Add(Exchange("A rifle is a weapon."));
        AiHistory.Add(new AiExchange
        {
            WhenUtc = new DateTime(2026, 10, 20, 12, 5, 0, DateTimeKind.Utc),
            Provider = AiProviders.OpenAi,
            Model = "gpt-4o-mini",
            Transcript = "where is the crate",
            Error = "The provider answered 401: bad key"
        });

        var model = new AiHistoryViewModel();
        var window = new AiHistoryWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var texts = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>()
                .Select(text => text.Text).ToList();
            Assert.IsTrue(texts.Contains("A rifle is a weapon."), "The answer is listed.");
            Assert.IsTrue(texts.Contains("The provider answered 401: bad key"), "The reason it failed is listed.");
            Assert.IsTrue(texts.Contains("what is a rifle"), "The question is listed.");
            Assert.IsTrue(texts.Contains("gpt-4o-mini"), "The model is listed.");
            Assert.AreEqual(2, model.Rows.Count, "Both exchanges are listed.");
            Assert.IsTrue(model.HasEntries, "There is something to clear.");
            Assert.IsTrue(model.Rows[0].Failed, "The newest exchange is the one that failed.");
            Assert.AreEqual("The provider answered 401: bad key", model.Rows[0].CopyText, "A click copies the reason.");
        }
        finally
        {
            window.Close();
            model.Dispose();
            AiHistory.Use(null);
        }
        return true;
    }, default);
}
