using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.MiniMap;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.MiniMap;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.MiniMap;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MiniMapWindow, MiniMapSettingsOverlay and BtnToggleMiniMap_Click: the mini-map, the settings that
/// decide how it looks, and the place it is left on the screen.
///
/// The settings and the place live in the reference's own cache files, so the WPF app and this one show the same
/// mini-map and put it in the same spot.
/// </summary>
[TestClass]
public sealed class MiniMapWindowTests
{
    /// <summary>A cache folder of its own, so a test never touches the real one.</summary>
    private sealed class Cache : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-minimap-window-" + Guid.NewGuid().ToString("N"));

        public Cache() => Directory.CreateDirectory(Folder);

        public string PathFor(string name) => Path.Combine(Folder, name);

        /// <summary>
        /// The settings are written in the background, so the folder can still be in use when the test is over; the
        /// delete waits for the write that is in flight rather than failing over it.
        /// </summary>
        public void Dispose()
        {
            for (int attempt = 0; attempt < 20; attempt++)
            {
                try
                {
                    if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
                    return;
                }
                catch (IOException) { Thread.Sleep(50); }
                catch (UnauthorizedAccessException) { Thread.Sleep(50); }
            }
        }
    }

    private static MiniMapViewModel Model(Cache cache, MapViewModel? map = null) =>
        new(new LegacyMiniMapSettingsStore(cache.Folder), new LegacyMiniMapPositionStore(cache.Folder)) { Map = map };

    // ---- the window ---------------------------------------------------------------------------

    [TestMethod]
    public async Task TheMiniMapOpensAndShowsTheMap() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var cache = new Cache();
        var map = new MapViewModel();
        map.Apply(MapTests.SyntheticMap());
        var model = Model(cache, map);
        var window = new MiniMapWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var surface = window.GetVisualDescendants().OfType<MapSurface>().Single();
            Assert.IsNotNull(surface.Image, "The mini-map shows the map the workspace shows.");
            Assert.AreSame(map.Image, surface.Image, "It is the same picture, not a copy of it.");
            Assert.IsTrue(window.IsVisible, "The mini-map is a window of its own.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-minimap.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheShapeAndTheSizeFollowTheSettings() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var cache = new Cache();
        var model = Model(cache);
        await model.LoadAsync();
        var window = new MiniMapWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();
            var shape = window.FindControl<Border>("Shape")!;

            Assert.AreEqual(260.0, window.Width, 0.01, "The reference's default size.");
            Assert.AreEqual(260.0, window.Height, 0.01, "A circle is as tall as it is wide.");
            Assert.AreEqual(130.0, shape.CornerRadius.TopLeft, 0.01, "The circle's radius is half its size.");

            model.ShapeIndex = MiniMapShapes.Widescreen;
            model.Size = 320;
            window.UpdateLayout();

            Assert.AreEqual(320.0, window.Width, 0.01, "The window follows the size the settings say.");
            Assert.AreEqual(180.0, window.Height, 0.01, "The 16:9 shape is shorter than it is wide.");
            Assert.AreEqual(12.0, shape.CornerRadius.TopLeft, 0.01, "The 16:9 shape keeps the reference's small radius.");

            model.ShapeIndex = MiniMapShapes.Square;
            model.Size = 1000;
            window.UpdateLayout();

            Assert.AreEqual(800.0, window.Width, 0.01, "The size is held at the largest the reference allows.");
            Assert.AreEqual(12.0, shape.CornerRadius.TopLeft, 0.01, "The square keeps the small radius too.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheTimeAndPopulationPillsFollowTheSettings() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var cache = new Cache();
        var model = Model(cache);
        await model.LoadAsync();
        model.ServerTime = "14:30";
        model.ServerPlayers = "42 / 200";
        var window = new MiniMapWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();
            var time = window.FindControl<Border>("TimePill")!;
            var pop = window.FindControl<Border>("PopPill")!;

            Assert.IsFalse(time.IsEffectivelyVisible, "The clock is hidden until it is asked for.");
            Assert.IsFalse(pop.IsEffectivelyVisible, "The population is hidden until it is asked for.");

            model.ShowTime = true;
            model.ShowPop = true;
            window.UpdateLayout();

            Assert.IsTrue(time.IsEffectivelyVisible, "The clock shows when the settings say so.");
            Assert.IsTrue(pop.IsEffectivelyVisible, "The population shows when the settings say so.");
            Assert.AreEqual("14:30", window.FindControl<TextBlock>("TimeText")!.Text, "The pill shows the server's clock.");
            Assert.AreEqual("42 / 200", window.FindControl<TextBlock>("PopText")!.Text, "The pill shows the population.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the settings panel --------------------------------------------------------------------

    [TestMethod]
    public async Task TheSettingsPanelOpensBesideTheMap() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var cache = new Cache();
        var model = Model(cache);
        await model.LoadAsync();
        var window = new MiniMapWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();
            var popup = window.FindControl<global::Avalonia.Controls.Primitives.Popup>("SettingsPopup")!;
            Assert.IsFalse(popup.IsOpen, "The settings are closed until the gear is pressed.");

            var gear = window.FindControl<Button>("SettingsButton")!;
            gear.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            window.UpdateLayout();

            Assert.IsTrue(popup.IsOpen, "The gear opens the settings.");
            var panel = popup.Child as MiniMapSettingsPanel;
            Assert.IsNotNull(panel, "The settings live in the popup, where the size slider cannot drag them away.");
            Assert.IsNotNull(panel!.GetVisualDescendants().OfType<ComboBox>()
                .Single(combo => global::Avalonia.Automation.AutomationProperties.GetName(combo) == "Mini-map shape"),
                "The panel offers the shape the reference offers.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the settings -------------------------------------------------------------------------

    [TestMethod]
    public async Task TheSettingsRoundTripThroughTheStore() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var cache = new Cache();
        var model = Model(cache);
        await model.LoadAsync();

        Assert.AreEqual(MiniMapShapes.Circle, model.ShapeIndex, "A mini-map with no settings is a circle.");
        Assert.AreEqual(260.0, model.Size, 0.01, "It is the reference's default size.");
        Assert.AreEqual(1.0, model.Opacity, 0.01, "It is fully opaque.");
        Assert.IsFalse(model.ShowTime, "The clock is off until it is asked for.");
        Assert.IsFalse(model.ShowPop, "The population is off until it is asked for.");

        model.ShapeIndex = MiniMapShapes.Square;
        model.Size = 300;
        model.Opacity = 0.5;
        model.ShowTime = true;
        model.ShowPop = true;
        model.ShowGrid = false;
        await model.SaveAsync();

        var read = Model(cache);
        await read.LoadAsync();
        Assert.AreEqual(MiniMapShapes.Square, read.ShapeIndex, "The shape is the one that was written.");
        Assert.AreEqual(300.0, read.Size, 0.01, "The size is the one that was written.");
        Assert.AreEqual(0.5, read.Opacity, 0.01, "The opacity is the one that was written.");
        Assert.IsTrue(read.ShowTime, "The clock is still shown.");
        Assert.IsTrue(read.ShowPop, "The population is still shown.");
        Assert.IsFalse(read.ShowGrid, "The grid is still hidden.");

        // The file the reference writes is the one that was read: the settings are shared, not copied.
        Assert.IsTrue(File.Exists(cache.PathFor("minimap_settings.json")),
            "The settings live in the reference's own cache file.");
    }, default);

    // ---- the place ----------------------------------------------------------------------------

    [TestMethod]
    public async Task ThePlaceIsSavedAndRestored() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var cache = new Cache();
        var model = Model(cache);
        await model.LoadAsync();
        var window = new MiniMapWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();
            window.Position = new global::Avalonia.PixelPoint(640, 360);

            // The place is written when the mini-map is moved.
            var store = new LegacyMiniMapPositionStore(cache.Folder);
            MiniMapPosition? place = null;
            await CompanionListenerViewModelTests.Eventually(() =>
            {
                place = store.LoadAsync(CancellationToken.None).GetAwaiter().GetResult();
                return place is { Left: 640, Top: 360 };
            });
            Assert.IsNotNull(place, "The place the mini-map was moved to is written.");
            Assert.AreEqual(640, place!.Left, "The left edge is the one the mini-map was moved to.");
            Assert.AreEqual(360, place.Top, "The top edge is the one the mini-map was moved to.");

            // A mini-map that is opened again comes back where it was left.
            var again = Model(cache);
            await again.LoadAsync();
            Assert.IsNotNull(again.SavedPosition, "The place is read back.");
            var restored = new MiniMapWindow { DataContext = again };
            try
            {
                restored.Show();
                restored.UpdateLayout();
                Assert.AreEqual(640, restored.Position.X, "The mini-map comes back at the left edge it was left at.");
                Assert.AreEqual(360, restored.Position.Y, "The mini-map comes back at the top edge it was left at.");
            }
            finally { restored.Close(); }
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the shell opens it once --------------------------------------------------------------

    private sealed class NoServers : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    [TestMethod]
    public async Task ASecondRequestFocusesTheOpenMiniMap() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var cache = new Cache();
        var model = new RustPlusDesk.Avalonia.Features.Servers.ServersViewModel(
            new NoServers(), new NoPicker(), "profiles.json", new UnusedSessionFactory(),
            miniMapSettings: new LegacyMiniMapSettingsStore(cache.Folder),
            miniMapPosition: new LegacyMiniMapPositionStore(cache.Folder));
        var window = new MainWindow { DataContext = model };
        try
        {
            window.Show();

            // Reference BtnMiniMap: the button on the map page asks the shell to open the mini-map.
            var workspace = window.GetVisualDescendants().OfType<MapWorkspace>().Single();
            var button = workspace.GetVisualDescendants().OfType<Button>()
                .Single(control => global::Avalonia.Automation.AutomationProperties.GetName(control) == "Mini-map");
            Assert.AreSame(model.OpenMiniMapCommand, button.Command, "The button on the map page asks for the mini-map.");

            button.Command!.Execute(null);
            await CompanionListenerViewModelTests.Eventually(() => window.MiniMapWindow is not null);
            var mini = window.MiniMapWindow!;
            Assert.IsTrue(mini.IsVisible, "The shell opens the mini-map in a window of its own.");

            button.Command!.Execute(null);
            Assert.AreSame(mini, window.MiniMapWindow, "Asking again does not open a second one.");
            Assert.AreEqual(1, mini.FocusCount, "The mini-map that is open is brought forward instead.");

            // The mini-map is not owned by the main window, so it has to be closed by hand or the app never stops.
            window.Close();
            Assert.IsFalse(mini.IsVisible, "The mini-map goes down with the app.");
        }
        finally
        {
            window.MiniMapWindow?.Close();
            window.Close();
        }
        return true;
    }, default);
}
