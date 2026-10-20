using System.Text.Json;
using Avalonia.Headless;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Players;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Core.Settings;
using RustPlusDesk.Core.Tracking;
using RustPlusDesk.Infrastructure.Tracking;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MainWindow.Players and TrackingService: the players the user keeps an eye on. The reference collected
/// how long a player stays on a server; that half was removed, so what is left is the bookmark: the players the user
/// added, the groups they are in and the order the groups are in.
/// </summary>
[TestClass]
public sealed class TrackedPlayerTests
{
    private static readonly DateTime Now = new(2026, 10, 20, 12, 0, 0, DateTimeKind.Utc);

    private static TrackedPlayer Player(string id, string name, string server = "Test Server", string group = "", bool online = false) =>
        new(id, name, server, group, "", online
            ? [new PlayerSession(Now.AddHours(-2), null)]
            : [new PlayerSession(Now.AddDays(-3), Now.AddDays(-3).AddHours(1))], IsBMOnly: true);

    // ---- the rules -----------------------------------------------------------------------------

    [TestMethod]
    public void TheFilterKeepsThePlayerWhoseNameOrServerMatches()
    {
        var players = new[] { Player("a", "Alice"), Player("b", "Bob", "Other Server") };
        Assert.AreEqual(1, TrackedPlayerRules.Filter(players, "ali").Count, "The name matches, whatever the case.");
        Assert.AreEqual(1, TrackedPlayerRules.Filter(players, "other").Count, "The server matches too.");
        Assert.AreEqual(0, TrackedPlayerRules.Filter(players, "nobody").Count);
        Assert.AreEqual(2, TrackedPlayerRules.Filter(players, "").Count, "No filter keeps everything.");
    }

    [TestMethod]
    public void ThePlayersAreGroupedByTheServerAndThenByTheirGroup()
    {
        var players = new[]
        {
            Player("a", "Alice", "Test Server", "Rivals"),
            Player("b", "Bob", "Test Server"),
            Player("c", "Carol", "Other Server"),
        };
        var servers = TrackedPlayerRules.Group(players, _ => []);
        Assert.AreEqual(2, servers.Count, "Two servers.");
        Assert.AreEqual("Test Server", servers[0].Name);
        Assert.AreEqual(2, servers[0].Groups.Count, "A group and the ungrouped ones.");
        Assert.AreEqual("Rivals", servers[0].Groups[0].GroupName);
        Assert.AreEqual(TrackedPlayerRules.UngroupedName, servers[0].Groups[1].GroupName,
            "The players without a group come after the ones that have one.");
    }

    [TestMethod]
    public void TheGroupsComeInTheOrderTheUserPutThemIn()
    {
        var players = new[]
        {
            Player("a", "Alice", "Test Server", "Rivals"),
            Player("b", "Bob", "Test Server", "Friends"),
        };
        var servers = TrackedPlayerRules.Group(players, _ => ["Friends", "Rivals"]);
        Assert.AreEqual("Friends", servers[0].Groups[0].GroupName, "The order the user chose comes first.");
        Assert.AreEqual("Rivals", servers[0].Groups[1].GroupName);

        var unordered = TrackedPlayerRules.Group(players, _ => []);
        Assert.AreEqual("Friends", unordered[0].Groups[0].GroupName, "Without an order the groups are in their name.");
    }

    [TestMethod]
    public void APlayerWithoutAServerIsListedUnderTheLegacyName()
    {
        var servers = TrackedPlayerRules.Group([Player("a", "Alice", "")], _ => []);
        Assert.AreEqual(TrackedPlayerRules.LegacyServerName, servers[0].Name);
    }

    [TestMethod]
    public void ThePlayTimeIsShownAsHoursAndMinutesAndOnlyWhileThePlayerIsOn()
    {
        Assert.AreEqual("02:30", TrackedPlayerRules.PlayTimeText(TimeSpan.FromMinutes(150)));
        Assert.AreEqual("00:05", TrackedPlayerRules.PlayTimeText(TimeSpan.FromMinutes(5)));
        Assert.AreEqual("", Player("a", "Alice").PlayTimeStr(Now), "A player who is not on has no play time.");
        Assert.AreEqual("02:00", Player("a", "Alice", online: true).PlayTimeStr(Now), "A player who is on is timed.");
    }

    [TestMethod]
    public void TheSessionsOlderThanTwelveWeeksAreDropped()
    {
        var player = new TrackedPlayer("a", "Alice", "Test Server", "", "",
            [new PlayerSession(Now.AddDays(-100), Now.AddDays(-100)), new PlayerSession(Now.AddDays(-1), null)], true);
        var pruned = TrackedPlayerRules.Prune(player, Now.AddDays(-TrackedPlayerRules.SessionRetentionDays));
        Assert.AreEqual(1, pruned.Sessions.Count, "Only the recent session is kept.");
    }

    [TestMethod]
    public void ASteamProfileIdIsSeventeenDigitsThatStartWithSevenSixFiveSix()
    {
        Assert.IsTrue(TrackedPlayerRules.IsSteamProfileId("76561198000000001"));
        Assert.IsFalse(TrackedPlayerRules.IsSteamProfileId("Alice"));
        Assert.IsFalse(TrackedPlayerRules.IsSteamProfileId("7656119800000000"), "Sixteen digits is not one.");
    }

    // ---- the file ------------------------------------------------------------------------------

    [TestMethod]
    public async Task ThePlayersAreKeptInTheReferenceFileAndReadBack()
    {
        using var file = new TempFile("tracked_players.json");
        var store = new TrackedPlayerFile(file.Path);
        var players = new[] { Player("a", "Alice", "Test Server", "Rivals") };
        await store.SaveAsync(players, CancellationToken.None);

        var read = await store.LoadAsync(CancellationToken.None);
        Assert.AreEqual(1, read.Count);
        Assert.AreEqual("Alice", read[0].Name);
        Assert.AreEqual("Rivals", read[0].GroupName);
        Assert.AreEqual("Test Server", read[0].LastServerName);

        // The reference keeps the players in a list, and the names are the reference's own.
        using var document = JsonDocument.Parse(File.ReadAllText(file.Path));
        var root = document.RootElement;
        Assert.AreEqual(JsonValueKind.Array, root.ValueKind, "The file is the list the reference writes.");
        var first = root[0];
        Assert.AreEqual("a", first.GetProperty("BMId").GetString());
        Assert.AreEqual("Alice", first.GetProperty("Name").GetString());
        Assert.IsTrue(first.TryGetProperty("IsBMOnly", out _), "The reference's own field is kept.");
        Assert.IsFalse(first.TryGetProperty("IsOnline", out _), "The online flag is not part of the file.");
    }

    [TestMethod]
    public async Task AFileThatCannotBeReadIsNoPlayersRatherThanABrokenApp()
    {
        using var file = new TempFile("tracked_players.json");
        await File.WriteAllTextAsync(file.Path, "{ not json");
        var store = new TrackedPlayerFile(file.Path);
        Assert.AreEqual(0, (await store.LoadAsync(CancellationToken.None)).Count);
    }

    [TestMethod]
    public async Task TheSessionsOlderThanTwelveWeeksAreNotWritten()
    {
        using var file = new TempFile("tracked_players.json");
        var store = new TrackedPlayerFile(file.Path);
        var player = new TrackedPlayer("a", "Alice", "Test Server", "", "",
            [new PlayerSession(DateTime.UtcNow.AddDays(-100), DateTime.UtcNow.AddDays(-100))], true);
        await store.SaveAsync([player], CancellationToken.None);

        using var document = JsonDocument.Parse(File.ReadAllText(file.Path));
        Assert.AreEqual(0, document.RootElement[0].GetProperty("Sessions").GetArrayLength(),
            "The old session is not written.");
    }

    // ---- the workspace -------------------------------------------------------------------------

    [TestMethod]
    public async Task TheTrackedPlayersAreShownAndCanBeChanged()
    {
        using var file = new TempFile("tracked_players.json");
        var settings = new ShopNoticeTests.MemorySettings();
        var store = new TrackedPlayerFile(file.Path);
        await store.SaveAsync([Player("a", "Alice", "Test Server", "Rivals")], CancellationToken.None);

        var model = new PlayersViewModel(store, settings);
        await model.LoadAsync(CancellationToken.None);

        Assert.AreEqual(1, model.Servers.Count, "The server the players were seen on.");
        Assert.AreEqual("Test Server", model.Servers[0].Name);
        Assert.AreEqual("Rivals", model.Servers[0].Groups[0].GroupName);
        Assert.AreEqual("Alice", model.Servers[0].Groups[0].Players[0].Name);

        // Reference BtnRenameTracked_Click: the player is given another name.
        var row = model.Servers[0].Groups[0].Players[0];
        row.Name = "Alice the Second";
        await model.RenameCommand.ExecuteAsync(row);
        Assert.AreEqual("Alice the Second", (await store.LoadAsync(CancellationToken.None))[0].Name);

        // Reference BtnGroupTracked_Click: the player is put in another group.
        row.GroupName = "Friends";
        await model.SetGroupCommand.ExecuteAsync(row);
        var groups = model.Servers[0].Groups.Select(group => group.GroupName).ToArray();
        Assert.IsTrue(groups.Contains("Friends"), "The new group is there.");
        Assert.IsFalse(groups.Contains("Rivals"), "The old group is not.");

        // Reference BtnRemoveTracked_Click: the player is no longer tracked.
        await model.RemoveCommand.ExecuteAsync(row);
        Assert.AreEqual(0, (await store.LoadAsync(CancellationToken.None)).Count);
        Assert.AreEqual(0, model.Servers.Count);
    }

    [TestMethod]
    public async Task TheGroupIsFoldedAwayAndTheChoiceIsKept()
    {
        using var file = new TempFile("tracked_players.json");
        var settings = new ShopNoticeTests.MemorySettings();
        var store = new TrackedPlayerFile(file.Path);
        await store.SaveAsync([Player("a", "Alice", "Test Server", "Rivals")], CancellationToken.None);

        var model = new PlayersViewModel(store, settings);
        await model.LoadAsync(CancellationToken.None);
        var group = model.Servers[0].Groups[0];
        Assert.IsTrue(group.IsExpanded, "A group is open unless the user folded it away.");

        model.ToggleGroupCommand.Execute(group);
        Assert.IsFalse(model.Servers[0].Groups[0].IsExpanded, "The group is folded away.");
        Assert.IsFalse(settings.Get(SettingKeys.GroupStates, new Dictionary<string, bool>())["Test Server|Rivals"],
            "The choice is kept in the settings.");

        // The group opens again.
        model.ToggleGroupCommand.Execute(model.Servers[0].Groups[0]);
        Assert.IsTrue(model.Servers[0].Groups[0].IsExpanded);
    }

    [TestMethod]
    public async Task TheGroupsArePutInTheOrderTheUserWantsThemIn()
    {
        using var file = new TempFile("tracked_players.json");
        var settings = new ShopNoticeTests.MemorySettings();
        var store = new TrackedPlayerFile(file.Path);
        await store.SaveAsync([Player("a", "Alice", "Test Server", "Rivals"), Player("b", "Bob", "Test Server", "Friends")],
            CancellationToken.None);

        var model = new PlayersViewModel(store, settings);
        await model.LoadAsync(CancellationToken.None);
        Assert.AreEqual("Friends", model.Servers[0].Groups[0].GroupName, "Without an order the groups are in their name.");

        model.MoveGroupDownCommand.Execute(model.Servers[0].Groups[0]);
        Assert.AreEqual("Rivals", model.Servers[0].Groups[0].GroupName, "The group was moved down.");
        CollectionAssert.AreEqual(new[] { "Rivals", "Friends" },
            settings.Get(SettingKeys.GroupOrder, new Dictionary<string, List<string>>())["Test Server"]);
    }

    [TestMethod]
    public async Task TheWorkspaceIsOpenedOnTheServerThatIsSelected()
    {
        using var file = new TempFile("tracked_players.json");
        var model = new PlayersViewModel(new TrackedPlayerFile(file.Path), new ShopNoticeTests.MemorySettings());
        model.Open("Test Server", connected: true);
        Assert.AreEqual(PlayersViewModel.TrackingRemovedMessage, model.Status,
            "The reference says the player tracking was removed.");
        model.Open(null, connected: false);
        Assert.AreNotEqual(PlayersViewModel.TrackingRemovedMessage, model.Status,
            "Without a server the reference says to connect first.");
    }

    [TestMethod]
    public async Task TheWorkspaceIsReachedFromTheRail() => await UiTestSession.Current.Dispatch(async () =>
    {
        var model = new ServersViewModel(new Infrastructure.Profiles.LegacyServerCatalog(),
            new Picker(), "profiles.json", new UnusedSessionFactory(),
            appSettings: new ShopNoticeTests.MemorySettings());
        var window = new global::Avalonia.Controls.Window { Width = 1000, Height = 700, DataContext = model };
        try
        {
            Assert.IsFalse(model.HasPlayersTab, "The Players tab is opt-in.");
            window.Show();
            window.UpdateLayout();
            Assert.IsFalse(window.GetVisualDescendants().OfType<global::Avalonia.Controls.Button>()
                .Any(button => global::Avalonia.Automation.AutomationProperties.GetName(button) == "Players"),
                "The rail entry is not there until the user turns it on.");
        }
        finally { await model.ShutdownAsync(); window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheViewShowsTheTrackedPlayers() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var file = new TempFile("tracked_players.json");
        var store = new TrackedPlayerFile(file.Path);
        await store.SaveAsync([Player("a", "Alice", "Test Server", "Rivals")], CancellationToken.None);
        var model = new PlayersViewModel(store, new ShopNoticeTests.MemorySettings());
        await model.LoadAsync(CancellationToken.None);
        model.IsTrackedTab = true;

        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new global::Avalonia.Controls.Border
            {
                Padding = new global::Avalonia.Thickness(16),
                Child = new PlayersView { DataContext = model }
            }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            var texts = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>().Select(text => text.Text).ToList();
            Assert.IsTrue(texts.Contains("Alice"), "The player is listed.");
            Assert.IsTrue(texts.Contains("Rivals"), "The group is listed.");
            Assert.IsTrue(texts.Contains("Test Server"), "The server is listed.");
            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-players.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    /// <summary>A file in a folder of its own, removed when the test is over.</summary>
    private sealed class TempFile : IDisposable
    {
        public string Directory { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rustplus-players-" + Guid.NewGuid().ToString("N"));
        public string Path { get; }

        public TempFile(string name)
        {
            Path = System.IO.Path.Combine(Directory, name);
            System.IO.Directory.CreateDirectory(Directory);
        }

        public void Dispose() => System.IO.Directory.Delete(Directory, recursive: true);
    }

    // ---- the group manager ----------------------------------------------------------------------

    [TestMethod]
    public async Task TheGroupManagerPutsThePlayersInTheGroupAndTakesThemOutAgain()
    {
        using var file = new TempFile("tracked_players.json");
        var settings = new ShopNoticeTests.MemorySettings();
        var store = new TrackedPlayerFile(file.Path);
        await store.SaveAsync(
            [Player("a", "Alice", "Test Server", "Rivals"), Player("b", "Bob", "Test Server", "Rivals")],
            CancellationToken.None);

        var model = new PlayersViewModel(store, settings);
        await model.LoadAsync(CancellationToken.None);

        model.OpenGroupManagerCommand.Execute(null);
        Assert.IsTrue(model.IsManagingGroups, "The group manager is open.");
        Assert.AreEqual(2, model.GroupPlayers.Count, "Every player is listed.");

        // An existing group is picked: its name, its colour and its players come with it.
        model.ChooseGroupCommand.Execute("Rivals");
        Assert.AreEqual("Rivals", model.GroupName);
        Assert.IsTrue(model.GroupPlayers.All(row => row.IsSelected), "The players in the group are checked.");

        // Bob is taken out of the group: the players that are not checked leave the group that is being saved.
        model.GroupPlayers.Single(row => row.BMId == "b").IsSelected = false;
        await model.SaveGroupCommand.ExecuteAsync(null);

        var saved = await store.LoadAsync(CancellationToken.None);
        Assert.AreEqual("Rivals", saved.Single(player => player.BMId == "a").GroupName, "Alice stays in the group.");
        Assert.AreEqual("", saved.Single(player => player.BMId == "b").GroupName, "Bob is out of the group.");
        Assert.IsFalse(model.IsManagingGroups, "The manager is closed once the group is saved.");

        // A new name moves the players that are checked into a new group.
        model.OpenGroupManagerCommand.Execute(null);
        model.GroupPlayers.Single(row => row.BMId == "a").IsSelected = true;
        model.GroupName = "Friends";
        model.GroupColour = "Blue";
        await model.SaveGroupCommand.ExecuteAsync(null);

        saved = await store.LoadAsync(CancellationToken.None);
        Assert.AreEqual("Friends", saved.Single(player => player.BMId == "a").GroupName, "Alice is in the new group.");
        Assert.AreEqual("Blue", saved.Single(player => player.BMId == "a").GroupColor, "The group has its colour.");
    }

    [TestMethod]
    public async Task AGroupWithoutANameIsNotSaved()
    {
        using var file = new TempFile("tracked_players.json");
        var store = new TrackedPlayerFile(file.Path);
        await store.SaveAsync([Player("a", "Alice")], CancellationToken.None);

        var model = new PlayersViewModel(store, new ShopNoticeTests.MemorySettings());
        await model.LoadAsync(CancellationToken.None);
        model.OpenGroupManagerCommand.Execute(null);
        model.GroupName = "   ";
        await model.SaveGroupCommand.ExecuteAsync(null);

        Assert.IsTrue(model.IsManagingGroups, "The manager stays open.");
        Assert.AreEqual("Please enter a group name.", model.GroupError, "The reference says what is wrong.");
        Assert.AreEqual("", (await store.LoadAsync(CancellationToken.None))[0].GroupName, "Nothing was saved.");
    }

    /// <summary>The file picker the tests never use.</summary>
    private sealed class Picker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => throw new AssertFailedException("The file picker should not run in this test.");
    }
}
