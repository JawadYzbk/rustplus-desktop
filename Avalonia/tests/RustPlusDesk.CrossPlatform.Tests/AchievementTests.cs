using global::Avalonia.Headless;
using global::Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Achievements;
using RustPlusDesk.Core.Achievements;
using RustPlusDesk.Infrastructure.Achievements;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference AchievementService and AchievementCatalog: what has been earned, what has not been looked at, and the
/// artwork the list is cut from.
/// </summary>
[TestClass]
public sealed class AchievementTests
{
    private sealed class MemoryStore : IAchievementStore
    {
        public AchievementState State { get; set; } = new();
        public int Saves { get; private set; }
        public AchievementState Load() => State;
        public void Save(AchievementState state) { State = state; Saves++; }
    }

    [TestMethod]
    public void TheCatalog_HoldsEveryAchievement_WithItsSheetCell()
    {
        Assert.AreEqual(31, AchievementCatalog.All.Count);
        Assert.AreEqual("I'm smarter now", AchievementCatalog.All[0].Name);
        Assert.AreEqual(30, AchievementCatalog.All[30].SheetIndex);
        Assert.IsTrue(AchievementCatalog.All[30].IsSecret, "The console command stays secret until it is earned.");
        Assert.IsNotNull(AchievementCatalog.Find(AchievementCatalog.ConsoleCommand));
        Assert.IsNull(AchievementCatalog.Find("not-an-achievement"));

        // The sheet is eight cells wide with an eight pixel gap between them.
        Assert.AreEqual((0, 0), AchievementCatalog.CellFor(0));
        Assert.AreEqual((136, 0), AchievementCatalog.CellFor(1));
        Assert.AreEqual((0, 136), AchievementCatalog.CellFor(8));
        Assert.AreEqual((816, 408), AchievementCatalog.CellFor(30));
    }

    [TestMethod]
    public void Unlocking_RecordsTheAchievement_AndMarksItUnseen()
    {
        var store = new MemoryStore();
        var tracker = new AchievementTracker(store);
        tracker.Load();

        Assert.AreEqual(0, tracker.EarnedCount);
        Assert.IsFalse(tracker.IsEarned(AchievementCatalog.TeamMate));

        tracker.Unlock(AchievementCatalog.TeamMate);
        Assert.IsTrue(tracker.IsEarned(AchievementCatalog.TeamMate));
        Assert.IsTrue(tracker.IsUnseen(AchievementCatalog.TeamMate));
        Assert.AreEqual(1, tracker.EarnedCount);
        Assert.AreEqual(1, tracker.UnseenCount);
        Assert.IsNotNull(tracker.EarnedAt(AchievementCatalog.TeamMate));
    }

    [TestMethod]
    public void UnlockingTwice_ChangesNothing()
    {
        var store = new MemoryStore();
        var tracker = new AchievementTracker(store);
        tracker.Load();
        tracker.Unlock(AchievementCatalog.TeamMate);
        int saves = store.Saves;
        tracker.Unlock(AchievementCatalog.TeamMate);
        Assert.AreEqual(1, tracker.EarnedCount);
        Assert.AreEqual(saves, store.Saves, "The second call does not write again.");
    }

    [TestMethod]
    public void AnUnknownId_IsNotWritten()
    {
        var store = new MemoryStore();
        var tracker = new AchievementTracker(store);
        tracker.Load();
        tracker.Unlock("not-an-achievement");
        Assert.AreEqual(0, tracker.EarnedCount);
        Assert.AreEqual(0, store.Saves);
    }

    [TestMethod]
    public void TheCelebrationIsHeldBack_WhileStartupIsStillSettling()
    {
        var tracker = new AchievementTracker(new MemoryStore());
        tracker.Load();
        AchievementDef? announced = null;
        tracker.Unlocked += definition => announced = definition;

        tracker.SuppressUnlockedEvents = true;
        tracker.Unlock(AchievementCatalog.TeamMate);
        Assert.IsNull(announced, "Startup must not bury the app under toasts.");
        Assert.AreEqual(1, tracker.UnseenCount, "The badge still tells them there is something to look at.");

        tracker.SuppressUnlockedEvents = false;
        tracker.Unlock(AchievementCatalog.RaidCalculator);
        Assert.IsNotNull(announced);
        Assert.AreEqual("No GP wasted", announced!.Name);
    }

    [TestMethod]
    public void MarkingSeen_ClearsTheBadge()
    {
        var tracker = new AchievementTracker(new MemoryStore());
        tracker.Load();
        tracker.Unlock(AchievementCatalog.TeamMate);
        Assert.AreEqual(1, tracker.UnseenCount);
        tracker.MarkAllSeen();
        Assert.AreEqual(0, tracker.UnseenCount);
        Assert.IsTrue(tracker.IsEarned(AchievementCatalog.TeamMate), "The achievement itself is kept.");
    }

    [TestMethod]
    public void TheStateRoundTripsThroughTheFile()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-achievements-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, "achievements.json");
        try
        {
            var store = new AchievementFileStore(path);
            var tracker = new AchievementTracker(store);
            tracker.Load();
            tracker.Unlock(AchievementCatalog.TeamMate);
            tracker.Unlock(AchievementCatalog.Clan);

            var reloaded = new AchievementTracker(new AchievementFileStore(path));
            reloaded.Load();
            Assert.AreEqual(2, reloaded.EarnedCount);
            Assert.IsTrue(reloaded.IsEarned(AchievementCatalog.TeamMate));
            Assert.IsTrue(reloaded.IsUnseen(AchievementCatalog.Clan), "The unseen list is kept too.");
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public void AFileThatCannotBeRead_StartsAFreshSet()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-achievements-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, "achievements.json");
        File.WriteAllText(path, "not json at all");
        try
        {
            var tracker = new AchievementTracker(new AchievementFileStore(path));
            tracker.Load();
            Assert.AreEqual(0, tracker.EarnedCount);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public void TheList_ShowsTheSecretAchievementAsQuestionMarks_UntilItIsEarned()
    {
        var model = new AchievementsViewModel();
        model.Load();
        var secret = model.Rows.Single(row => row.Definition.Id == AchievementCatalog.ConsoleCommand);
        Assert.AreEqual("???", secret.Name);
        Assert.IsNull(secret.Icon, "A secret achievement shows no artwork before it is earned.");

        model.Unlock(AchievementCatalog.ConsoleCommand);
        var earned = model.Rows.Single(row => row.Definition.Id == AchievementCatalog.ConsoleCommand);
        Assert.AreEqual("Console Command", earned.Name);
        Assert.AreEqual(1, model.EarnedCount);
        Assert.AreEqual("1 of 31 earned", model.Summary);
    }

    [TestMethod]
    public async Task TheList_RendersTheArtworkOfAnEarnedAchievement() => await UiTestSession.Current.Dispatch(async () =>
    {
        var model = new AchievementsViewModel();
        model.Load();
        model.Unlock(AchievementCatalog.TeamMate);

        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new RustPlusDesk.Avalonia.Features.Achievements.AchievementsView { DataContext = model }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            var images = window.GetVisualDescendants().OfType<global::Avalonia.Controls.Image>().ToList();
            Assert.AreEqual(1, images.Count(image => image.Source is not null), "Only the earned achievement shows artwork.");
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-achievements.png"), global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public void TheLastAchievementRaisesTheCompletedEvent()
    {
        var model = new AchievementsViewModel();
        model.Load();
        var completed = 0;
        model.Completed += () => completed++;
        foreach (var definition in RustPlusDesk.Core.Achievements.AchievementCatalog.All) model.Unlock(definition.Id);
        Assert.AreEqual(1, completed, "The question is asked once, when the set is complete.");
        Assert.IsTrue(model.IsComplete);
    }
}
