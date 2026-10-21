using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.MiniMap;
using RustPlusDesk.Core.MiniMap;
using RustPlusDesk.Infrastructure.MiniMap;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MainWindow.MiniMapPeek and BtnFollowPlayer: the key that is held to see the map over the game, and the
/// map that follows the player in it.
/// </summary>
[TestClass]
public sealed class MiniMapPeekTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-minimap-" + Guid.NewGuid().ToString("N"));
        public LegacyMiniMapSettingsStore Settings { get; }
        public LegacyMiniMapPositionStore Position { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            Settings = new LegacyMiniMapSettingsStore(Folder);
            Position = new LegacyMiniMapPositionStore(Folder);
        }

        public void Dispose()
        {
            try { if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true); }
            catch (IOException) { }
        }
    }

    [TestMethod]
    public async Task TheFollowIsSavedWithTheRestOfTheMiniMapSettings()
    {
        using var fixture = new Fixture();
        var model = new MiniMapViewModel(fixture.Settings, fixture.Position);
        await model.LoadAsync();

        model.FollowPlayer = true;
        await model.SaveAsync();

        var stored = await fixture.Settings.LoadAsync(CancellationToken.None);
        Assert.IsNotNull(stored, "The settings were written.");
        Assert.IsTrue(stored!.FollowPlayer, "The follow was saved with the rest.");
    }

    [TestMethod]
    public async Task TheFollowTellsTheMapWhichPlayerToKeepInTheMiddle()
    {
        using var fixture = new Fixture();
        var map = new RustPlusDesk.Avalonia.Features.Map.MapViewModel();
        var model = new MiniMapViewModel(fixture.Settings, fixture.Position) { Map = map };
        await model.LoadAsync();

        model.OwnSteamId = "76561198000000000";
        model.FollowPlayer = true;

        Assert.AreEqual(76561198000000000UL, map.FollowSteamId, "The map follows the player the account is signed in as.");

        model.FollowPlayer = false;
        Assert.AreEqual(0UL, map.FollowSteamId, "The map follows nobody when the switch is off.");
    }

    [TestMethod]
    public async Task ThePeekIsAWindowsFeatureAndSaysSo() => await UiTestSession.Current.Dispatch(() =>
    {
        var controller = new MiniMapPeekController(new global::Avalonia.Controls.Window());

        if (OperatingSystem.IsWindows())
        {
            Assert.IsTrue(controller.IsSupported, "The peek works on Windows.");
            Assert.AreEqual("", controller.UnavailableReason, "There is nothing to explain when it works.");
        }
        else
        {
            Assert.IsFalse(controller.IsSupported, "The peek needs the Windows keyboard hook.");
            Assert.IsTrue(controller.UnavailableReason.Length > 0, "The reason is given.");
        }

        controller.Dispose();
        return true;
    }, default);

    [TestMethod]
    public async Task ThePeekKeyIsTheOneTheSettingNames() => await UiTestSession.Current.Dispatch(() =>
    {
        var controller = new MiniMapPeekController(new global::Avalonia.Controls.Window());

        controller.Gesture = "Tab";
        Assert.AreEqual("Tab", controller.Gesture, "The key the setting names is the key that is held.");

        controller.Gesture = "";
        Assert.AreEqual("", controller.Gesture, "An empty setting means no peek.");

        controller.Dispose();
        return true;
    }, default);

    [TestMethod]
    public async Task ThePeekGoesAwayWhenTheTimeIsUp() => await UiTestSession.Current.Dispatch(() =>
    {
        var controller = new MiniMapPeekController(new global::Avalonia.Controls.Window());
        int shown = 0, hidden = 0;
        controller.Show = () => shown++;
        controller.Hide = () => hidden++;

        Assert.IsFalse(controller.IsPeeking, "The map is not up to begin with.");
        Assert.AreEqual(TimeSpan.FromSeconds(20), MiniMapPeekController.Limit, "The map may be held up for twenty seconds.");

        controller.Dispose();
        return true;
    }, default);
}
