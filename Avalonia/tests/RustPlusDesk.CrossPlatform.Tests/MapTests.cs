using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class MapTests
{
    [TestMethod]
    public void Geometry_MatchesReferenceWorldRectAndRoundTrips()
    {
        // Reference ComputeWorldRectFromWorldSize: a 4500 world in a 2000 px image with 2000 padding.
        var geometry = new MapGeometry(2000, 2000, 4500);
        var (x, y, side) = geometry.WorldRect;
        Assert.AreEqual(2000 * 4500 / 6500.0, side, 1e-9);
        Assert.AreEqual((2000 - side) / 2, x, 1e-9);
        Assert.AreEqual(x, y, 1e-9);
        // World (0, size) is the top-left corner of the playable square; (size, 0) the bottom-right.
        var topLeft = geometry.WorldToImage(0, 4500);
        Assert.AreEqual(x, topLeft.X, 1e-9);
        Assert.AreEqual(y, topLeft.Y, 1e-9);
        var bottomRight = geometry.WorldToImage(4500, 0);
        Assert.AreEqual(x + side, bottomRight.X, 1e-9);
        Assert.AreEqual(y + side, bottomRight.Y, 1e-9);
        var world = geometry.ImageToWorld(geometry.WorldToImage(1234, 3210).X, geometry.WorldToImage(1234, 3210).Y);
        Assert.AreEqual(1234, world.X, 1e-6);
        Assert.AreEqual(3210, world.Y, 1e-6);
        // Points beyond the padding clamp like the reference.
        Assert.AreEqual(geometry.WorldToImage(-1000, 0), geometry.WorldToImage(-99999, 0));
    }

    [TestMethod]
    public void Grid_UsesReferenceCellCountsAndLabels()
    {
        Assert.AreEqual(29, MapGrid.CellCount(4250));
        Assert.AreEqual(4250 / 29.0, MapGrid.CellSize(4250), 1e-9);
        Assert.AreEqual(30, MapGrid.CellCount(4500));
        Assert.AreEqual("A", MapGrid.ColumnLabel(0));
        Assert.AreEqual("Z", MapGrid.ColumnLabel(25));
        Assert.AreEqual("AA", MapGrid.ColumnLabel(26));
        Assert.AreEqual("AC", MapGrid.ColumnLabel(28));
        Assert.AreEqual("A0", MapGrid.Label(4500, 0, 4500));
        Assert.AreEqual("AD29", MapGrid.Label(4500, 4499, 1));
        // The reference regression: a 4250 map must not drift (E24 was once reported as D23).
        double cell = MapGrid.CellSize(4250);
        Assert.AreEqual("E24", MapGrid.Label(4250, 4.5 * cell, 4250 - 24.5 * cell));
        Assert.AreEqual("off-grid", MapGrid.Label(0, 1, 1));
    }

    [TestMethod]
    public void Monuments_UseReferenceNamesVariantsAndIcons()
    {
        Assert.AreEqual("Large Oil Rig", MonumentCatalog.DisplayName("large_oil_rig"));
        Assert.AreEqual("oilrig.png", MonumentCatalog.IconFile("oil_rig_small"));
        Assert.AreEqual("Sulfur Quarry", MonumentCatalog.DisplayName("mining_quarry_sulfur_display_name"));
        Assert.AreEqual("sulfurquarry.png", MonumentCatalog.IconFile("mining_quarry_sulfur_display_name"));
        Assert.AreEqual("Underwater Labs", MonumentCatalog.DisplayName("underwater_lab_b"));
        Assert.AreEqual("underwater.png", MonumentCatalog.IconFile("underwater_lab_b"));
        Assert.AreEqual("Fishing Village A", MonumentCatalog.DisplayName("fishing_village_a"));
        Assert.AreEqual("Train Tunnel (A)", MonumentCatalog.DisplayName("Train Tunnel A"));
        Assert.AreEqual(0.6, MonumentCatalog.IconScale("swamp"));
        Assert.IsTrue(MonumentCatalog.IsTrainTunnel("train_tunnel_display_name"));
        Assert.IsNull(MonumentCatalog.IconFile("some_modded_thing"));
    }

    [TestMethod]
    public async Task MapViewModel_LoadsRetriesAndResets()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new MapViewModel();
            var session = new MapSession(SyntheticMap()) { FailuresBeforeSuccess = 1 };
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(20));
            await model.RunAsync(session, cts.Token);
            Assert.AreEqual(2, session.Calls);
            Assert.IsTrue(model.HasMap);
            Assert.AreEqual(4500, model.Geometry!.WorldSize);
            Assert.AreEqual(3, model.Monuments.Count);
            Assert.AreEqual("", model.Message);
            Assert.IsFalse(model.IsLoading);
            model.Reset();
            Assert.IsFalse(model.HasMap);
            Assert.IsNull(model.Geometry);
            return true;
        }, default);
    }

    [TestMethod]
    public async Task MapSurface_RendersFitsZoomsPansAndHoversMonuments()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new MapViewModel();
            model.Apply(SyntheticMap());
            var surface = new MapSurface
            {
                Image = model.Image, Geometry = model.Geometry, Monuments = model.Monuments, ResetRequest = model.ViewResetRequest
            };
            var window = new Window { Width = 900, Height = 700, Content = surface };
            try
            {
                window.Show();
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);
                double fit = surface.Scale;
                Assert.AreEqual(Math.Min(900 / 1024.0, 700 / 1024.0), fit, 1e-9);
                var centre = new Point(450, 350);
                window.MouseWheel(centre, new Vector(0, 1));
                Assert.AreEqual(fit * 1.25, surface.Scale, 1e-9);
                // The image point under the pointer stays put while zooming.
                Assert.AreEqual(centre.X, surface.ImageToScreen(512, 512).X, 1.0);
                var before = surface.Offset;
                window.MouseDown(new Point(300, 300), MouseButton.Left);
                window.MouseMove(new Point(340, 320), RawInputModifiers.LeftMouseButton);
                window.MouseUp(new Point(340, 320), MouseButton.Left);
                Assert.AreEqual(before + new Vector(40, 20), surface.Offset);

                surface.FitToView();
                var rig = model.Monuments.First(monument => monument.Token == "large_oil_rig");
                var (ix, iy) = model.Geometry!.WorldToImage(rig.X, rig.Y);
                var onScreen = surface.ImageToScreen(ix, iy);
                window.MouseMove(onScreen);
                Assert.AreSame(rig, surface.Hovered);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map.png"), PngBitmapEncoderOptions.Default);
                }
                surface.ShowMonuments = false;
                window.MouseMove(onScreen + new Vector(1, 0));
                Assert.IsNull(surface.Hovered);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }

    /// <summary>A 1024 px synthetic map image (sea with an island) so tests never use real server data.</summary>
    internal static ServerMap SyntheticMap()
    {
        var target = new RenderTargetBitmap(new PixelSize(1024, 1024));
        using (var context = target.CreateDrawingContext())
        {
            context.FillRectangle(new SolidColorBrush(Color.FromRgb(18, 64, 82)), new Rect(0, 0, 1024, 1024));
            context.DrawEllipse(new SolidColorBrush(Color.FromRgb(112, 138, 84)), null, new Point(512, 512), 330, 300);
        }
        using var stream = new MemoryStream();
        target.Save(stream, PngBitmapEncoderOptions.Default);
        return new ServerMap(stream.ToArray(), 1024, 1024, 500, 4500,
            [new("large_oil_rig", 4300, 300), new("mining_quarry_sulfur_display_name", 2000, 2500), new("fishing_village_a", 900, 3800)]);
    }

    internal static IReadOnlyList<MapMarker> SyntheticMarkers() =>
    [
        new(1, MapMarkerKind.Player, 2200, 2300) { Name = "Teammate", SteamId = 76561198000000001 },
        new(2, MapMarkerKind.VendingMachine, 1500, 1500) { Name = "Scrap shop", Orders = [new(-932201673, 10, 1, 50, 3)] },
        new(3, MapMarkerKind.VendingMachine, 1600, 1500) { Name = "Empty", IsOutOfStock = true },
        new(4, MapMarkerKind.CargoShip, 300, 2000) { Rotation = 90 },
        new(5, MapMarkerKind.PatrolHelicopter, 3000, 3000) { Rotation = 45 },
        new(6, MapMarkerKind.GenericRadius, 2500, 2500) { Radius = 120, Color = 0xFFFF0000 },
        new(7, MapMarkerKind.Crate, 2600, 1200)
    ];

    [TestMethod]
    public async Task MapViewModel_PollsMarkersAfterTheMapLoads()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new MapViewModel { MarkerInterval = TimeSpan.FromMilliseconds(30) };
            var session = new MapSession(SyntheticMap()) { Markers = SyntheticMarkers() };
            using var cts = new CancellationTokenSource();
            var running = model.RunAsync(session, cts.Token);
            await CompanionListenerViewModelTests.Eventually(() => model.Markers.Count == 7 && session.MarkerCalls >= 2);
            session.FailMarkers = true;
            int before = session.MarkerCalls;
            await CompanionListenerViewModelTests.Eventually(() => session.MarkerCalls > before + 1);
            Assert.AreEqual(7, model.Markers.Count, "A failed poll keeps the last markers.");
            cts.Cancel();
            await running.WaitAsync(TimeSpan.FromSeconds(4));
            model.Reset();
            Assert.AreEqual(0, model.Markers.Count);
            return true;
        }, default);
    }

    [TestMethod]
    public async Task MapSurface_DrawsMarkersHoversThemAndRespectsToggles()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new MapViewModel();
            model.Apply(SyntheticMap());
            var surface = new MapSurface { Image = model.Image, Geometry = model.Geometry, Monuments = model.Monuments, Markers = SyntheticMarkers() };
            var window = new Window { Width = 900, Height = 700, Content = surface };
            try
            {
                window.Show();
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);
                Point At(MapMarker marker) { var (x, y) = model.Geometry!.WorldToImage(marker.X, marker.Y); return surface.ImageToScreen(x, y); }
                var shop = surface.Markers!.First(marker => marker.Name == "Scrap shop");
                window.MouseMove(At(shop));
                Assert.AreSame(shop, surface.HoveredMarker);
                StringAssert.StartsWith(MapSurface.Describe(shop, 4500), "Scrap shop · 1 offer · ");
                var player = surface.Markers!.First(marker => marker.Kind == MapMarkerKind.Player);
                window.MouseMove(At(player));
                Assert.AreSame(player, surface.HoveredMarker);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-markers.png"), PngBitmapEncoderOptions.Default);
                }
                surface.ShowPlayers = false;
                window.MouseMove(At(player) + new Vector(1, 0));
                Assert.IsNull(surface.HoveredMarker);
                surface.ShowShops = false;
                window.MouseMove(At(shop));
                Assert.IsNull(surface.HoveredMarker);
                Assert.IsNotNull(MapSurface.MarkerIcon("cargo.png"));
                Assert.IsNotNull(MapSurface.MarkerIcon("vending_orange.png"));
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }

    [TestMethod]
    public async Task CenterOnPlayer_CentresTheSurfaceOnTheirMarker_KeepingTheZoom()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new MapViewModel();
            Assert.IsFalse(model.CenterOnPlayer(76561198000000001), "Nothing to centre on before the map loads.");
            model.Apply(SyntheticMap());
            model.Markers = SyntheticMarkers();
            var surface = new MapSurface { Image = model.Image, Geometry = model.Geometry, Markers = model.Markers };
            model.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(MapViewModel.FocusRequest)) surface.FocusRequest = model.FocusRequest; };
            var window = new Window { Width = 900, Height = 700, Content = surface };
            try
            {
                window.Show();
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);
                surface.ZoomAt(new Point(100, 100), 2);
                double zoom = surface.Scale;
                Assert.IsFalse(model.CenterOnPlayer(42), "Unknown players are not on the map.");
                Assert.IsFalse(model.CenterOnPlayer(0));
                Assert.IsTrue(model.CenterOnPlayer(76561198000000001));
                var player = model.Markers.First(marker => marker.Kind == MapMarkerKind.Player);
                var (x, y) = model.Geometry!.WorldToImage(player.X, player.Y);
                var centre = surface.ImageToScreen(x, y);
                Assert.AreEqual(surface.Bounds.Width / 2, centre.X, 0.01);
                Assert.AreEqual(surface.Bounds.Height / 2, centre.Y, 0.01);
                Assert.AreEqual(zoom, surface.Scale);
                surface.ZoomAt(new Point(10, 10), 1.25);
                Assert.IsTrue(model.CenterOnPlayer(76561198000000001), "Repeating the request centres again.");
                Assert.AreEqual(surface.Bounds.Width / 2, surface.ImageToScreen(x, y).X, 0.01);
                model.Reset();
                Assert.IsNull(model.FocusRequest);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }

    private sealed class MapSession(ServerMap map) : IServerMarkerSession
    {
        public IReadOnlyList<MapMarker> Markers { get; init; } = [];
        public bool FailMarkers { get; set; }
        public int MarkerCalls { get; private set; }
        public Task<IReadOnlyList<MapMarker>> ReadMarkersAsync(CancellationToken ct)
        {
            MarkerCalls++;
            return FailMarkers ? Task.FromException<IReadOnlyList<MapMarker>>(new ServerConnectionException(ConnectionFailure.Unreachable)) : Task.FromResult(Markers);
        }
        public int FailuresBeforeSuccess { get; init; }
        public int Calls { get; private set; }
        public IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Synthetic", 0, 0, 0);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public Task<ServerMap> ReadMapAsync(CancellationToken ct)
            => ++Calls <= FailuresBeforeSuccess ? Task.FromException<ServerMap>(new ServerConnectionException(ConnectionFailure.Unreachable)) : Task.FromResult(map);
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
