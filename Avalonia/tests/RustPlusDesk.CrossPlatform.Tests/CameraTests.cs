using System.Text.Json;
using System.Text.Json.Nodes;
using Avalonia.Headless;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Cameras;
using RustPlusDesk.Core.Cameras;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MainWindow.Camera and ServerProfile.CameraIds/CameraNames: the cameras the user added for a server.
///
/// The cameras live in the profile, so the WPF app and this one show the same cameras; the picture of a tile is
/// grabbed from the camera itself, one camera at a time.
/// </summary>
[TestClass]
public sealed class CameraTests
{
    private const string Host = "10.0.0.1";
    private const int Port = 28082;
    private const string SteamId = "76561198000000001";

    private static string Profiles() => $"[{Host}:{Port}]";

    private static string ProfileKey() => RustPlusDesk.Infrastructure.Profiles.LegacyProfileIdentity.Create(Host, Port, SteamId);

    private sealed class Fixture : IDisposable
    {
        public string Directory { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "rustplus-cameras-" + Guid.NewGuid().ToString("N"));
        public string Path { get; }

        public Fixture(string contents = "")
        {
            Path = System.IO.Path.Combine(Directory, "profiles.json");
            System.IO.Directory.CreateDirectory(Directory);
            File.WriteAllText(Path, contents.Length > 0 ? contents :
                "[{\"Host\":\"" + Host + "\",\"Port\":" + Port + ",\"SteamId64\":\"" + SteamId +
                "\",\"PlayerToken\":\"token\",\"CameraIds\":[\"cam-1\",\"cam-2\"],\"CameraNames\":{\"cam-1\":\"Gate camera\"}}]");
        }

        public void Dispose()
        {
            if (System.IO.Directory.Exists(Directory)) System.IO.Directory.Delete(Directory, recursive: true);
        }
    }

    private sealed class Snapshots : ICameraSnapshotSource
    {
        /// <summary>A one-pixel PNG, so the tile has a real picture to show.</summary>
        public static readonly byte[] Picture =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82,
        ];

        public List<string> Asked { get; } = new();

        public Task<byte[]?> GrabAsync(string profilesPath, string profileKey, string cameraId, CancellationToken ct)
        {
            Asked.Add(cameraId);
            return Task.FromResult<byte[]?>(Picture);
        }
    }

    // ---- the profile ---------------------------------------------------------------------------

    [TestMethod]
    public async Task TheCamerasAreReadFromTheProfile()
    {
        using var file = new Fixture();
        var store = new LegacyCameraStore();
        var cameras = await store.LoadAsync(file.Path, ProfileKey(), CancellationToken.None);

        Assert.AreEqual(2, cameras.Count, "Both cameras are read.");
        Assert.AreEqual("cam-1", cameras[0].Id);
        Assert.AreEqual("Gate camera", cameras[0].Name);
        Assert.AreEqual("cam-2", cameras[1].Id);
        Assert.AreEqual("cam-2", cameras[1].DisplayName, "A camera without a name is shown by its identifier.");
    }

    [TestMethod]
    public async Task TheCamerasAreWrittenBackToTheProfile()
    {
        using var file = new Fixture();
        var store = new LegacyCameraStore();
        await store.SaveAsync(file.Path, ProfileKey(),
            [new CameraEntry("cam-3", "Roof camera"), new CameraEntry("cam-4", "")], CancellationToken.None);

        using var document = JsonDocument.Parse(await File.ReadAllTextAsync(file.Path));
        var profile = document.RootElement[0];
        var ids = profile.GetProperty("CameraIds").EnumerateArray().Select(value => value.GetString()).ToArray();
        CollectionAssert.AreEqual(new[] { "cam-3", "cam-4" }, ids, "The cameras are the reference's own list.");
        Assert.AreEqual("Roof camera", profile.GetProperty("CameraNames").GetProperty("cam-3").GetString());
        Assert.IsFalse(profile.GetProperty("CameraNames").TryGetProperty("cam-4", out _), "A camera without a name is not named.");

        var read = await store.LoadAsync(file.Path, ProfileKey(), CancellationToken.None);
        Assert.AreEqual(2, read.Count, "They are read back.");
    }

    [TestMethod]
    public async Task AProfileThatIsNotThereHasNoCameras()
    {
        using var file = new Fixture("""[{"Host":"other.test","Port":28082,"SteamId64":"76561198000000002"}]""");
        var store = new LegacyCameraStore();
        Assert.AreEqual(0, (await store.LoadAsync(file.Path, ProfileKey(), CancellationToken.None)).Count);
    }

    // ---- the workspace -------------------------------------------------------------------------

    [TestMethod]
    public async Task TheCamerasAreShownAndChanged()
    {
        using var file = new Fixture();
        var store = new LegacyCameraStore();
        var snapshots = new Snapshots();
        var model = new CamerasViewModel(store, snapshots);

        await model.LoadAsync(file.Path, ProfileKey());
        Assert.AreEqual(2, model.Rows.Count, "Both cameras are listed.");
        Assert.AreEqual("Gate camera", model.Rows[0].Name);
        Assert.IsFalse(model.IsEmpty, "There is something to show.");

        // Reference CommitAddCamera: the camera the user typed in is added.
        model.NewId = "cam-9";
        model.NewName = "New camera";
        await model.AddCommand.ExecuteAsync(null);
        Assert.AreEqual(3, model.Rows.Count, "The camera was added.");
        var saved = await store.LoadAsync(file.Path, ProfileKey(), CancellationToken.None);
        Assert.AreEqual("cam-9", saved[^1].Id);

        // Reference CommitDeleteCamera: the camera is no longer added.
        await model.RemoveCommand.ExecuteAsync(model.Rows[0]);
        Assert.AreEqual(2, model.Rows.Count, "The camera was removed.");
        saved = await store.LoadAsync(file.Path, ProfileKey(), CancellationToken.None);
        Assert.IsFalse(saved.Any(camera => camera.Id == "cam-1"), "It is gone from the profile too.");
    }

    [TestMethod]
    public async Task ACameraWithoutAnIdentifierIsNotAdded()
    {
        using var file = new Fixture();
        var model = new CamerasViewModel(new LegacyCameraStore(), new Snapshots());
        await model.LoadAsync(file.Path, ProfileKey());

        model.NewId = "   ";
        await model.AddCommand.ExecuteAsync(null);
        Assert.AreEqual("Please enter a camera identifier.", model.Error, "The reference says what is wrong.");
        Assert.AreEqual(2, model.Rows.Count, "Nothing was added.");
    }

    [TestMethod]
    public async Task ACameraThatIsAddedTwiceIsRefused()
    {
        using var file = new Fixture();
        var model = new CamerasViewModel(new LegacyCameraStore(), new Snapshots());
        await model.LoadAsync(file.Path, ProfileKey());

        model.NewId = "CAM-1";
        await model.AddCommand.ExecuteAsync(null);
        Assert.AreEqual("That camera is already added.", model.Error);
        Assert.AreEqual(2, model.Rows.Count);
    }

    [TestMethod]
    public async Task TheTileShowsWhatTheCameraSees() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var file = new Fixture();
        var snapshots = new Snapshots();
        var model = new CamerasViewModel(new LegacyCameraStore(), snapshots);
        await model.LoadAsync(file.Path, ProfileKey());

        await model.GrabNextAsync();
        CollectionAssert.AreEqual(new[] { "cam-1" }, snapshots.Asked, "One camera at a time is asked.");
        Assert.IsTrue(model.Rows[0].IsLive, "The camera answered.");
        Assert.IsNotNull(model.Rows[0].Picture, "The tile shows the picture.");
        return true;
    }, default);

    [TestMethod]
    public async Task TheViewShowsTheCameras() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var file = new Fixture();
        var model = new CamerasViewModel(new LegacyCameraStore(), new Snapshots());
        await model.LoadAsync(file.Path, ProfileKey());

        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new global::Avalonia.Controls.Border
            {
                Padding = new global::Avalonia.Thickness(16),
                Child = new CamerasView { DataContext = model }
            }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            var texts = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>().Select(text => text.Text).ToList();
            Assert.IsTrue(texts.Contains("Gate camera"), "The camera is listed.");
            Assert.IsFalse(window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>()
                .Any(text => text.Text == "No cameras found" && text.IsEffectivelyVisible), "The empty state is gone when there are cameras.");
            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-cameras.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheEmptyStateIsShownWhenThereAreNoCameras() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var file = new Fixture("""[{"Host":"10.0.0.1","Port":28082,"SteamId64":"76561198000000001"}]""");
        var model = new CamerasViewModel(new LegacyCameraStore(), new Snapshots());
        await model.LoadAsync(file.Path, ProfileKey());

        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new global::Avalonia.Controls.Border
            {
                Padding = new global::Avalonia.Thickness(16),
                Child = new CamerasView { DataContext = model }
            }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            var texts = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>().Select(text => text.Text).ToList();
            Assert.IsTrue(texts.Contains("No cameras found"), "The reference's empty state.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the window opens from the tile ---------------------------------------------------------

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    [TestMethod]
    public void OpeningACameraAsksTheShellForAWindowOfItsOwn()
    {
        var model = new RustPlusDesk.Avalonia.Features.Servers.ServersViewModel(
            new NoServers(), new NoPicker(), "profiles.json", new UnusedSessionFactory(),
            cameraStore: new LegacyCameraStore(), cameraSnapshots: new Snapshots(), cameraSessions: new FakeSessions());

        CameraRow? asked = null;
        model.CameraOpenRequested += (_, row) => asked = row;

        var page = model.Cameras!;
        page.OpenCommand.Execute(new CameraRow(new CameraEntry("cam-1", "Gate camera")));

        Assert.IsNotNull(asked, "The shell is asked for a window of its own.");
        Assert.AreEqual("cam-1", asked!.Id, "It is the camera the user clicked.");
        Assert.AreEqual("Gate camera", asked.Name, "It keeps the name the user gave it.");
        Assert.IsNotNull(model.CameraSessions, "The app has the camera sessions the window needs.");
    }

    private sealed class NoServers : RustPlusDesk.Core.Servers.IServerCatalog
    {
        public Task<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>>([]);
    }

    // ---- the camera window ----------------------------------------------------------------------

    private sealed class FakeSession : ICameraSession
    {
        public string CameraId { get; } = "cam-1";
        public CameraInfo Info { get; } = new("cam-1", "Gate camera", CameraKind.Ptz, 160, 90,
            CameraControlFlags.Movement | CameraControlFlags.Mouse | CameraControlFlags.Fire | CameraControlFlags.Reload | CameraControlFlags.Crosshair);
        public int TargetFps { get; set; }
        public event Action<byte[]>? FrameRendered;
        public event Action<IReadOnlyList<CameraEntity>, double>? EntitiesUpdated;
        public event Action<string>? KeepAliveFailed;

        public Task SendInputAsync(CameraButtons buttons, float mouseDeltaX, float mouseDeltaY, CancellationToken ct) => Task.CompletedTask;
        public Task LookAsync(float deltaX, float deltaY, CancellationToken ct) => Task.CompletedTask;
        public Task ShootAsync(CancellationToken ct) => Task.CompletedTask;
        public Task ReloadAsync(CancellationToken ct) => Task.CompletedTask;
        public Task ZoomAsync(CancellationToken ct) => Task.CompletedTask;
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;

        /// <summary>What the camera sends when it is asked for.</summary>
        public void Send()
        {
            FrameRendered?.Invoke(Snapshots.Picture);
            EntitiesUpdated?.Invoke([new CameraEntity(CameraEntityKind.Player, "Alice", 0, 0, 10)], 65);
        }

        /// <summary>What the camera says when it is gone.</summary>
        public void Fail(string message) => KeepAliveFailed?.Invoke(message);
    }

    private sealed class FakeSessions : ICameraSessionFactory
    {
        public FakeSession Session { get; } = new();

        public Task<ICameraSession> OpenAsync(string profilesPath, string profileKey, string cameraId, CancellationToken ct) =>
            Task.FromResult<ICameraSession>(Session);
    }

    [TestMethod]
    public async Task TheCameraWindowShowsTheLivePictureAndTheControls() => await UiTestSession.Current.Dispatch(async () =>
    {
        var sessions = new FakeSessions();
        var window = new CameraWindow(sessions, "profiles.json", "key", "cam-1", "Gate camera",
            () => new HashSet<string>(["Alice"], StringComparer.OrdinalIgnoreCase));
        try
        {
            window.Show();
            window.UpdateLayout();
            sessions.Session.Send();
            await Task.Delay(50);
            window.UpdateLayout();

            var texts = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>().Select(text => text.Text).ToList();
            Assert.IsTrue(texts.Contains("Gate camera"), "The camera's name is in the title.");
            Assert.IsTrue(window.GetVisualDescendants().OfType<global::Avalonia.Controls.Button>()
                .Any(button => global::Avalonia.Automation.AutomationProperties.GetName(button) == "Fire"), "The fire button is there.");

            var names = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>()
                .Where(text => text.Text == "Alice").ToList();
            Assert.IsTrue(names.Count > 0, "The player the camera sees is named in the picture.");
            var colour = (names[0].Foreground as global::Avalonia.Media.SolidColorBrush)?.Color;
            Assert.AreEqual(global::Avalonia.Media.Color.FromRgb(50, 205, 50), colour,
                "The reference draws the team players green.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-camera-window.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    // ---- the stick and the overlay ---------------------------------------------------------------

    [TestMethod]
    public void TheStickIsClampedToItsCircle()
    {
        var (x, y) = CameraInput.Stick(30, 0, 59);
        Assert.AreEqual(30.0 / 59.0, x, 0.001, "The stick keeps its place inside the circle.");
        Assert.AreEqual(0.0, y, 0.001, "Up is positive.");

        var (farX, farY) = CameraInput.Stick(200, 0, 59);
        Assert.AreEqual(1.0, farX, 0.001, "A drag past the edge is clamped.");
        Assert.AreEqual(0.0, farY, 0.001);
    }

    [TestMethod]
    public void TheStickResponseIsFineNearTheCentreAndFastAtTheEdge()
    {
        Assert.AreEqual(0f, CameraInput.Response(0.05, 16), "Below the threshold the camera is not moved.");
        Assert.AreEqual(16f, CameraInput.Response(1, 16), 0.001, "At the edge the camera moves at the full speed.");
        Assert.AreEqual(-4f, CameraInput.Response(-0.5, 16), 0.001, "The curve is quadratic.");
    }

    [TestMethod]
    public void TheStickSendsTheKeysOfItsDirection()
    {
        Assert.AreEqual(CameraButtons.Forward, CameraInput.MovementButtons(0, 1), "Up is forward.");
        Assert.AreEqual(CameraButtons.Backward, CameraInput.MovementButtons(0, -1), "Down is backward.");
        Assert.AreEqual(CameraButtons.Right | CameraButtons.Forward, CameraInput.MovementButtons(1, 1), "Both at once.");
        Assert.AreEqual(CameraButtons.None, CameraInput.MovementButtons(0.2, 0.2), "Inside the dead zone the camera is not moved.");
    }

    [TestMethod]
    public void APlayerIsDrawnWhereTheCameraSeesIt()
    {
        var ahead = new CameraEntity(CameraEntityKind.Player, "Alice", 0, 0, 10);
        var place = CameraOverlay.Place(ahead, 65, 160, 90);
        Assert.IsNotNull(place, "A player in front is drawn.");
        Assert.AreEqual(80.0, place!.Value.X, 0.001, "Straight ahead is the middle of the picture.");
        Assert.AreEqual(45.0, place.Value.Y, 0.001, "Straight ahead is the middle of the picture.");

        var behind = new CameraEntity(CameraEntityKind.Player, "Bob", 0, 0, -10);
        Assert.IsNull(CameraOverlay.Place(behind, 65, 160, 90), "A player behind the camera is not drawn.");

        var aside = new CameraEntity(CameraEntityKind.Player, "Carol", 1000, 0, 10);
        Assert.IsNull(CameraOverlay.Place(aside, 65, 160, 90), "A player outside the picture is not drawn.");
    }
}
