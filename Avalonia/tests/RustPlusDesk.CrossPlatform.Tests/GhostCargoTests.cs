using Avalonia.Controls;
using Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference dynamic-marker persistence (MainWindow.Map.Markers <c>_dynStates</c> and the missing-marker loop): a moving
/// event marker that stops appearing keeps its element alive for 9 polls, moving one velocity step per poll, and is
/// forgotten after 10. The predicted (ghost) cargo positions are fed to the cargo tracker with the reference
/// <c>isGhost</c> rules - they keep a ship alive but never dock it or announce.
/// </summary>
[TestClass]
public sealed class GhostCargoTests
{
    private const double World = 4500;
    private static readonly DateTime T0 = new(2026, 5, 1, 12, 0, 0, DateTimeKind.Utc);
    private static readonly MapMonument Harbor = new("harbor_display_name", 1000, 1000);
    private static readonly CargoAlertFlags All = new(true, true, true, true, true);
    private static MapMarker Cargo(ulong id, double x, double y) => new(id, MapMarkerKind.CargoShip, x, y);

    private static IReadOnlyList<MapMarker> Observe(PredictedMarkers predicted, params MapMarker[] markers) => predicted.Observe(markers);

    [TestMethod]
    public void Ghost_MovesAtTheLastVelocityThenIsForgotten()
    {
        var predicted = new PredictedMarkers();
        Assert.AreEqual(1, Observe(predicted, Cargo(1, 1000, 1000)).Count, "A live marker is drawn as is.");
        Assert.AreEqual(0, predicted.Ghosts.Count);
        Assert.AreEqual(1, Observe(predicted, Cargo(1, 1010, 1020)).Count, "The second real poll sets the velocity.");

        var ghost = Observe(predicted, []);
        Assert.AreEqual(1, ghost.Count, "The marker stays alive while it is missing.");
        Assert.AreEqual(1020, ghost[0].X, 0.001, "One velocity step per missing poll.");
        Assert.AreEqual(1040, ghost[0].Y, 0.001);
        Assert.AreEqual(MapMarkerKind.CargoShip, ghost[0].Kind);
        Assert.AreEqual(1ul, ghost[0].Id);
        Assert.AreEqual(1, predicted.Ghosts.Count, "The ghost is what the cargo tracker is fed.");

        for (int missing = 2; missing <= 9; missing++)
        {
            ghost = Observe(predicted, []);
            Assert.AreEqual(1, ghost.Count, $"Still alive on missing poll {missing}.");
            Assert.AreEqual(1010 + 10 * missing, ghost[0].X, 0.001);
        }
        Assert.AreEqual(0, Observe(predicted, []).Count, "The tenth missing poll forgets the marker.");
        Assert.AreEqual(0, predicted.Ghosts.Count);
    }

    [TestMethod]
    public void Ghost_KeepsAParkedMarkerInPlace()
    {
        var predicted = new PredictedMarkers();
        Observe(predicted, Cargo(1, 2000, 2000));
        var ghost = Observe(predicted, []);
        Assert.AreEqual(2000, ghost[0].X, 0.001, "A marker seen once has no velocity yet.");
        Assert.AreEqual(2000, ghost[0].Y, 0.001);
    }

    [TestMethod]
    public void Ghost_ExcludesPlayersAndStaticMarkers()
    {
        var predicted = new PredictedMarkers();
        Observe(predicted, new MapMarker(1, MapMarkerKind.Player, 100, 100), new MapMarker(2, MapMarkerKind.Crate, 200, 200),
            new MapMarker(3, MapMarkerKind.Explosion, 300, 300), new MapMarker(4, MapMarkerKind.VendingMachine, 400, 400));
        Assert.AreEqual(0, Observe(predicted, []).Count, "Only the moving event markers are predicted.");
        Assert.AreEqual(0, predicted.Ghosts.Count);
    }

    [TestMethod]
    public void Ghost_AHelicopterAndAVendorArePredictedToo()
    {
        var predicted = new PredictedMarkers();
        Observe(predicted, new MapMarker(1, MapMarkerKind.PatrolHelicopter, 1000, 1000) { Name = "Patrol Helicopter" },
            new MapMarker(2, MapMarkerKind.TravellingVendor, 2000, 2000));
        Observe(predicted, new MapMarker(1, MapMarkerKind.PatrolHelicopter, 1010, 1000) { Name = "Patrol Helicopter" },
            new MapMarker(2, MapMarkerKind.TravellingVendor, 2000, 2020));
        var ghost = Observe(predicted, []);
        Assert.AreEqual(2, ghost.Count);
        Assert.AreEqual(1020, ghost[0].X, 0.001);
        Assert.AreEqual("Patrol Helicopter", ghost[0].Name, "The ghost keeps the marker's name for the label.");
        Assert.AreEqual(2040, ghost[1].Y, 0.001, "The vendor is predicted on its own axis: its last 20 m step.");
    }

    [TestMethod]
    public void Ghost_TheRealMarkerReturningStopsThePrediction()
    {
        var predicted = new PredictedMarkers();
        Observe(predicted, Cargo(1, 1000, 1000));
        Observe(predicted, Cargo(1, 1010, 1010));
        Observe(predicted, []);
        var back = Observe(predicted, Cargo(1, 1050, 1050));
        Assert.AreEqual(1, back.Count);
        Assert.AreEqual(1050, back[0].X, 0.001, "The returned marker is drawn at its real position.");
        Assert.AreEqual(0, predicted.Ghosts.Count, "No ghost while the marker is live again.");
    }

    [TestMethod]
    public void Ghost_ResetForgetsEveryMarker()
    {
        var predicted = new PredictedMarkers();
        Observe(predicted, Cargo(1, 1000, 1000));
        Observe(predicted, Cargo(1, 1010, 1010));
        predicted.Reset();
        Assert.AreEqual(0, Observe(predicted, []).Count, "A new connection has no ghosts.");
        Assert.AreEqual(0, predicted.Ghosts.Count);
    }

    [TestMethod]
    public void Tracker_AGhostKeepsADockedShipDocked_AndTracksItsPosition()
    {
        var tracker = new CargoTracker(new MemoryCargo());
        var now = T0;
        // Park the ship at the harbour so it docks for real, then announce it after the 5 second delay.
        tracker.Observe([Cargo(1, 2000, 2000)], World, [Harbor], "10.0.0.1", All, now);
        now = now.AddSeconds(10);
        tracker.Observe([Cargo(1, 1010, 1010)], World, [Harbor], "10.0.0.1", All, now);
        now = now.AddSeconds(10);
        tracker.Observe([Cargo(1, 1010, 1010)], World, [Harbor], "10.0.0.1", All, now);
        var docked = tracker.Ships.Single();
        Assert.IsTrue(docked.IsDocked, "Parked at the harbour: docked.");
        DateTime? dockTime = docked.DockTime;

        // Ghosts: a far-away predicted position must not look like a departure.
        now = now.AddSeconds(10);
        tracker.ObserveGhosts([Cargo(1, 4000, 4000)], World, [Harbor], "10.0.0.1", All, now);
        var afterGhost = tracker.Ships.Single();
        Assert.IsTrue(afterGhost.IsDocked, "A ghost cannot undock the ship.");
        Assert.AreEqual(dockTime, afterGhost.DockTime, "A ghost cannot reset DockTime.");
        Assert.AreEqual(1010, afterGhost.LastX, 0.001, "A docked ship's ghost stays at its last position.");

        // A live moving marker still departs.
        now = now.AddSeconds(10);
        tracker.Observe([Cargo(1, 1500, 1010)], World, [Harbor], "10.0.0.1", All, now);
        Assert.IsFalse(tracker.Ships.Single().IsDocked, "The real marker departs.");
    }

    [TestMethod]
    public void Tracker_GhostsNeverCreateAShipAndKeepItAlive()
    {
        var tracker = new CargoTracker(new MemoryCargo());
        var now = T0;
        tracker.ObserveGhosts([Cargo(9, 1000, 1000)], World, [Harbor], "10.0.0.1", All, now);
        Assert.AreEqual(0, tracker.Ships.Count, "A ghost never creates a ship state.");

        tracker.Observe([Cargo(1, 2000, 2000)], World, [Harbor], "10.0.0.1", All, now);
        now = now.AddSeconds(10);
        tracker.Observe([Cargo(1, 2100, 2100)], World, [Harbor], "10.0.0.1", All, now);
        now = now.AddSeconds(10);
        tracker.ObserveGhosts([Cargo(1, 2200, 2200)], World, [Harbor], "10.0.0.1", All, now);
        Assert.AreEqual(2200, tracker.Ships.Single().LastX, 0.001, "The ghost position is taken as the last position.");

        // The 60 second despawn is measured from the last sighting, ghost or real.
        now = now.AddSeconds(59);
        tracker.Observe([], World, [Harbor], "10.0.0.1", All, now);
        Assert.AreEqual(1, tracker.Ships.Count, "Still within the despawn grace period.");
        now = now.AddSeconds(2);
        tracker.Observe([], World, [Harbor], "10.0.0.1", All, now);
        Assert.AreEqual(0, tracker.Ships.Count, "Gone 60 seconds after the last sighting.");
    }

    [TestMethod]
    public void Tracker_AGhostDoesNotAnnounceTheDocking()
    {
        var tracker = new CargoTracker(new MemoryCargo());
        var now = T0;
        tracker.Observe([Cargo(1, 2000, 2000)], World, [Harbor], "10.0.0.1", All, now);
        now = now.AddSeconds(10);
        tracker.Observe([Cargo(1, 1010, 1010)], World, [Harbor], "10.0.0.1", All, now);
        Assert.IsFalse(tracker.Ships.Single().IsDocked, "Sailing in, not docked yet.");
        // Only ghosts from here on: the ship parks at the harbour, but a ghost can never dock or announce it.
        now = now.AddSeconds(10);
        tracker.ObserveGhosts([Cargo(1, 1010, 1010)], World, [Harbor], "10.0.0.1", All, now);
        now = now.AddSeconds(10);
        tracker.ObserveGhosts([Cargo(1, 1010, 1010)], World, [Harbor], "10.0.0.1", All, now);
        Assert.IsFalse(tracker.Ships.Single().IsDocked, "Ghost markers cannot trigger docking.");
    }

    [TestMethod]
    public void MapViewModel_DrawsTheGhostsAndExposesThemForTheTracker()
    {
        var map = new MapViewModel();
        map.Markers = [Cargo(1, 1000, 1000)];
        Assert.AreEqual(1, map.DrawnMarkers.Count);
        Assert.AreEqual(0, map.GhostMarkers.Count);

        map.Markers = [Cargo(1, 1010, 1010)];
        map.Markers = [];
        Assert.AreEqual(1, map.DrawnMarkers.Count, "The surface keeps drawing the missing cargo.");
        Assert.AreEqual(1, map.GhostMarkers.Count, "The cargo tracker is fed the ghost.");
        Assert.AreEqual(1020, map.DrawnMarkers[0].X, 0.001);

        map.Reset();
        Assert.AreEqual(0, map.DrawnMarkers.Count, "A new connection clears the predicted markers.");
        Assert.AreEqual(0, map.GhostMarkers.Count);
    }

    [TestMethod]
    public async Task Map_DrawsTheGhostOverTheWorkspace() => await UiTestSession.Current.Dispatch(async () =>
    {
        var shell = new ServersViewModel(new NoServers(), new NoPicker(), "profiles.json", new UnusedSessionFactory());
        var workspace = new MapWorkspace { DataContext = shell };
        var window = new Window { Width = 1000, Height = 700, Content = workspace };
        try
        {
            window.Show();
            shell.Connection.Map.Markers = [Cargo(1, 1000, 1000)];
            shell.Connection.Map.Markers = [Cargo(1, 1010, 1010)];
            shell.Connection.Map.Markers = [];
            window.UpdateLayout();
            var surface = workspace.FindControl<MapSurface>("Surface")!;
            Assert.AreEqual(1, surface.Markers?.Count ?? 0, "The surface is bound to the predicted markers.");
            Assert.AreEqual(1020, surface.Markers![0].X, 0.001);
            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class MemoryCargo : ICargoMemory
    {
        public int Docking { get; private set; } = SettingsCargoMemory.DefaultDockingMinutes;
        public int FullLife { get; private set; }
        public Dictionary<string, (double X, double Y)> Triggers { get; } = [];

        public int DockingMinutes(string host) => Docking;
        public void SetDockingMinutes(string host, int minutes) => Docking = minutes;
        public (double X, double Y)? TriggerPoint(string host, string harbor) => Triggers.TryGetValue(harbor, out var point) ? point : null;
        public void SetTriggerPoint(string host, string harbor, double x, double y) => Triggers[harbor] = (x, y);
        public bool HasAnyTrigger(string host) => Triggers.Count > 0;
        public int FullLifeMinutes(string host) => FullLife;
        public void SetFullLife(string host, int minutes) => FullLife = minutes;
    }

    private sealed class NoServers : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([]);
    }

    private sealed class NoPicker : IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
