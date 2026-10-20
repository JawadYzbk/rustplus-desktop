using global::Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Infrastructure.Map;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference the map's overlay: the shape tools, the routes the user names and where the drawings are kept.
/// </summary>
[TestClass]
public sealed class MapDrawingTests
{
    // ---- the shape tools -------------------------------------------------------------------------

    [TestMethod]
    public void ALineRunsFromCornerToCorner()
    {
        var points = MapShapes.BuildShapePoints(MapToolMode.Line, new MapPoint(0, 0), new MapPoint(10, 5), false);
        Assert.AreEqual(2, points.Count);
        Assert.AreEqual(new MapPoint(0, 0), points[0]);
        Assert.AreEqual(new MapPoint(10, 5), points[1]);
    }

    [TestMethod]
    public void AnArrowIsOneContinuousStroke()
    {
        var points = MapShapes.BuildShapePoints(MapToolMode.Arrow, new MapPoint(0, 0), new MapPoint(10, 0), false);
        Assert.AreEqual(5, points.Count, "The shaft, one barb, the tip and the other barb.");
        Assert.AreEqual(new MapPoint(0, 0), points[0]);
        Assert.AreEqual(new MapPoint(10, 0), points[1]);
        Assert.AreEqual(new MapPoint(10, 0), points[3], "The stroke goes back to the tip.");
    }

    [TestMethod]
    public void ABoxIsClosed_AndACircleIsRound()
    {
        var box = MapShapes.BuildShapePoints(MapToolMode.Box, new MapPoint(0, 0), new MapPoint(10, 5), false);
        Assert.AreEqual(5, box.Count);
        Assert.AreEqual(box[0], box[4], "The box comes back to where it started.");

        var circle = MapShapes.BuildShapePoints(MapToolMode.Circle, new MapPoint(0, 0), new MapPoint(10, 10), false);
        Assert.AreEqual(49, circle.Count, "Forty-eight segments and the closing point.");
        Assert.AreEqual(new MapPoint(10, 5), circle[0], "The circle starts at the right of the ellipse.");
    }

    [TestMethod]
    public void ShiftSnapsTheDragToFortyFiveDegrees()
    {
        var snapped = MapShapes.SnapToAngle(new MapPoint(0, 0), new MapPoint(10, 1));
        Assert.AreEqual(10, Math.Round(snapped.X, 6), 6);
        Assert.AreEqual(0, Math.Round(snapped.Y, 6), 6);

        var points = MapShapes.BuildShapePoints(MapToolMode.Line, new MapPoint(0, 0), new MapPoint(10, 1), true);
        Assert.AreEqual(0, Math.Round(points[1].Y, 6), 6);
    }

    [TestMethod]
    public void ShiftMakesTheBoxSquare()
    {
        var points = MapShapes.BuildShapePoints(MapToolMode.Box, new MapPoint(0, 0), new MapPoint(10, 3), true);
        Assert.AreEqual(10, Math.Round(points[1].X, 6), 6);
        Assert.AreEqual(10, Math.Round(points[2].Y, 6), 6, "The box is as tall as it is wide.");
    }

    [TestMethod]
    public void AFreehandRouteIsThinned()
    {
        var points = new List<MapPoint>();
        for (int i = 0; i < 40; i++) points.Add(new MapPoint(i, i % 2 == 0 ? 0 : 0.01));
        var simplified = MapShapes.SimplifyPoints(points, 1.0);
        Assert.IsTrue(simplified.Count < points.Count, "The points on a straight line are dropped.");
        Assert.AreEqual(points[0], simplified[0]);
        Assert.AreEqual(points[^1], simplified[^1]);
    }

    // ---- the store ------------------------------------------------------------------------------

    [TestMethod]
    public void TheDrawingsRoundTripThroughTheStore()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-drawings-" + Guid.NewGuid().ToString("N"));
        try
        {
            var store = new MapDrawingStore(folder);
            var data = new MapOverlayData
            {
                Strokes = [new SavedStroke([new MapPoint(1, 2), new MapPoint(3, 4)], "#FF00FF00", 3)],
                Icons = [new SavedIcon("base.png", 10, 20, 30, "Base")],
                Texts = [new SavedText(5, 6, "Hello", "#FFFFFFFF", 12)],
                Routes = [new SavedRoute("r1", "Supply run", "#FF3FD7FF", 2, [new MapPoint(1, 1), new MapPoint(2, 2)], true, false)],
                WipeTimeUnix = 1700000000
            };
            store.Save("server.test:28082", "76561198000000001", data);

            var read = store.Load("server.test:28082", "76561198000000001");
            Assert.AreEqual(1, read.Strokes.Count);
            Assert.AreEqual(new MapPoint(1, 2), read.Strokes[0].Points[0]);
            Assert.AreEqual("#FF00FF00", read.Strokes[0].Color);
            Assert.AreEqual(1, read.Icons.Count);
            Assert.AreEqual("Base", read.Icons[0].Label);
            Assert.AreEqual(1, read.Texts.Count);
            Assert.AreEqual("Hello", read.Texts[0].Text);
            Assert.AreEqual(1, read.Routes.Count);
            Assert.AreEqual("Supply run", read.Routes[0].Name);
            Assert.IsTrue(read.Routes[0].Closed);
            Assert.IsFalse(read.Routes[0].Visible, "A hidden route stays hidden.");
            Assert.AreEqual(1700000000, read.WipeTimeUnix);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public void AFileThatCannotBeRead_LeavesTheMapWithoutDrawings()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-drawings-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(folder, "server.test_28082"));
        File.WriteAllText(Path.Combine(folder, "server.test_28082", "76561198000000001.json"), "not json");
        try
        {
            var read = new MapDrawingStore(folder).Load("server.test:28082", "76561198000000001");
            Assert.AreEqual(0, read.Strokes.Count);
            Assert.AreEqual(0, read.Routes.Count);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    // ---- the panel ------------------------------------------------------------------------------

    [TestMethod]
    public void ThePanel_DrawsShapes_NamesRoutes_AndUndoes()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-drawings-" + Guid.NewGuid().ToString("N"));
        try
        {
            var store = new MapDrawingStore(folder);
            var model = new MapDrawingsViewModel(store)
            {
                ServerKey = () => "server.test:28082",
                SteamId = () => "76561198000000001"
            };
            model.Load();

            model.ToolChoice = MapDrawingsViewModel.ToolChoices.First(choice => choice.Tool == MapToolMode.Line);
            model.DrawShape(new MapPoint(0, 0), new MapPoint(10, 10));
            Assert.AreEqual(1, model.Overlay.Strokes.Count);

            model.ToolChoice = MapDrawingsViewModel.ToolChoices.First(choice => choice.Tool == MapToolMode.Route);
            model.PlacePoint(new MapPoint(1, 1));
            model.PlacePoint(new MapPoint(2, 2));
            model.RouteName = "Supply run";
            model.FinishRouteCommand.Execute(null);
            Assert.AreEqual(1, model.Routes.Count);
            Assert.AreEqual("Supply run", model.Routes[0].Name);
            Assert.IsFalse(model.IsDrawing, "The route is finished.");

            model.UndoCommand.Execute(null);
            Assert.AreEqual(0, model.Overlay.Strokes.Count, "The last drawing is taken back.");

            model.DeleteRouteCommand.Execute(model.Routes[0]);
            Assert.AreEqual(0, model.Routes.Count);

            // What is on the map is what was written down.
            var written = store.Load("server.test:28082", "76561198000000001");
            Assert.AreEqual(0, written.Strokes.Count);
            Assert.AreEqual(0, written.Routes.Count);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public void ThePanel_ReadsTheDrawingsOfTheSelectedServer()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-drawings-" + Guid.NewGuid().ToString("N"));
        try
        {
            var store = new MapDrawingStore(folder);
            store.Save("server.test:28082", "76561198000000001", new MapOverlayData
            {
                Routes = [new SavedRoute("r1", "Supply run", "#FF3FD7FF", 2, [new MapPoint(1, 1), new MapPoint(2, 2)])]
            });

            var model = new MapDrawingsViewModel(store)
            {
                ServerKey = () => "server.test:28082",
                SteamId = () => "76561198000000001"
            };
            model.Load();
            Assert.AreEqual(1, model.Routes.Count);
            Assert.AreEqual("Supply run", model.Routes[0].Name);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }
    [TestMethod]
    public void TheIconAndTextTools_PlaceTheirThing()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-drawings-" + Guid.NewGuid().ToString("N"));
        try
        {
            var model = new MapDrawingsViewModel(new MapDrawingStore(folder))
            {
                ServerKey = () => "server.test_28082",
                SteamId = () => "76561198000000001"
            };
            model.Load();

            model.ToolChoice = MapDrawingsViewModel.ToolChoices.First(choice => choice.Tool == MapToolMode.Icon);
            model.PlacePoint(new MapPoint(10, 20));
            Assert.AreEqual(0, model.Overlay.Icons.Count, "No icon is picked yet.");
            model.IconChoice = model.Icons.First();
            model.PlacePoint(new MapPoint(10, 20));
            Assert.AreEqual(1, model.Overlay.Icons.Count);
            Assert.AreEqual(10, model.Overlay.Icons[0].X);

            model.ToolChoice = MapDrawingsViewModel.ToolChoices.First(choice => choice.Tool == MapToolMode.Text);
            model.TextValue = "Base camp";
            model.PlacePoint(new MapPoint(30, 40));
            Assert.AreEqual(1, model.Overlay.Texts.Count);
            Assert.AreEqual("Base camp", model.Overlay.Texts[0].Text);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public async Task TheDrawPanel_RendersWithItsTools() => await UiTestSession.Current.Dispatch(async () =>
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-drawings-" + Guid.NewGuid().ToString("N"));
        try
        {
            var model = new MapDrawingsViewModel(new MapDrawingStore(folder))
            {
                ServerKey = () => "server.test_28082",
                SteamId = () => "76561198000000001"
            };
            model.Load();
            model.ToolChoice = MapDrawingsViewModel.ToolChoices.First(choice => choice.Tool == MapToolMode.Line);
            model.DrawShape(new MapPoint(0, 0), new MapPoint(10, 10));

            var window = new global::Avalonia.Controls.Window
            {
                Width = 1000,
                Height = 700,
                Content = new RustPlusDesk.Avalonia.Shell.MapWorkspace
                {
                    DataContext = new RustPlusDesk.Avalonia.Features.Servers.ServersViewModel(
                        new LegacyServerCatalog(), new NoPicker(), "profiles.json", new UnusedSessionFactory())
                }
            };
            try
            {
                window.Show();
                window.UpdateLayout();
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-drawings.png"),
                        global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            }
            finally { window.Close(); }
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
        return true;
    }, default);

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
