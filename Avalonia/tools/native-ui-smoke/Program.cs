using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using RustPlusDesk.Avalonia;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.CrossPlatform.Tests;
using RustPlusDesk.Infrastructure.Items;

internal static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        // No desktop lifetime: production App initializes its exact styles without starting personal account services.
        AppBuilder.Configure<App>().UsePlatformDetect().SetupWithoutStarting();
        using var stop = new CancellationTokenSource();
        int result = 1;
        string output = Path.GetFullPath(args.FirstOrDefault() ?? Path.Combine(AppContext.BaseDirectory, "captures"));
        Directory.CreateDirectory(output);
        var factory = new FixtureSessions();
        var model = new ServersViewModel(new Catalog(), new Picker(), "fixture-profiles.json", factory, recyclerCatalog: new BundledRecyclerCatalog());
        var window = new MainWindow { DataContext = model, Title = "Rust+ Desktop — native UI smoke (fixtures only)" };
        window.Opened += async (_, _) =>
        {
            try
            {
                await model.RefreshCommand.ExecuteAsync(null);
                var connection = model.Connection.ConnectCommand.ExecuteAsync(null);
                await Task.Delay(200);
                Require(model.Connection.IsConnected, "Connected fixture state");
                Require(model.Connection.Map.TeamMembers.Count == 2, "Team fixture state");
                foreach (var theme in new[] { ThemeVariant.Dark, ThemeVariant.Light })
                foreach (var size in new[] { (1000, 700), (1380, 950), (1920, 1080) })
                {
                    window.Width = size.Item1; window.Height = size.Item2;
                    window.RequestedThemeVariant = theme;
                    await Task.Delay(160);
                    Capture(window, Path.Combine(output, $"shell-{theme}-{size.Item1}.png"));
                    var map = window.GetVisualDescendants().OfType<LiveMapView>().Single();
                    foreach (var button in map.GetVisualDescendants().OfType<Button>().Where(b => b.IsEffectivelyVisible))
                    {
                        var point = button.TranslatePoint(default, window)!.Value;
                        Require(point.X >= 0 && point.Y >= 0 && point.X + button.Bounds.Width <= window.ClientSize.Width + 1
                            && point.Y + button.Bounds.Height <= window.ClientSize.Height + 1, "Map action within client bounds");
                    }
                }
                window.Width = 1000; window.Height = 700;
                model.Connection.Map.SelectedMember = model.Connection.Map.TeamMembers[0];
                await Task.Delay(100);
                Require(window.GetVisualDescendants().OfType<MapSurface>().Single().ZoomLevel >= 4, "Native follow selection");
                Capture(window, Path.Combine(output, "team-follow.png"));
                model.Connection.Chat.Team.Draft = "Local fixture message";
                await model.Connection.Chat.Team.SendCommand.ExecuteAsync(null);
                Require(model.Connection.Chat.Team.Messages.Any(m => m.Text == "Local fixture message"), "Native chat bindings");
                model.ToggleSidebarCommand.Execute(null);
                await Task.Delay(80);
                Capture(window, Path.Combine(output, "sidebar-collapsed.png"));
                model.ShowDevicesCommand.Execute(null);
                model.OpenRecyclerCommand.Execute(null);
                await Task.Delay(100);
                Require(model.Recycler is not null && model.IsRecyclerSelected, "Recycler navigation");
                Capture(window, Path.Combine(output, "recycler.png"));
                model.CloseRecyclerCommand.Execute(null);
                await model.Connection.StopAsync();
                await connection;
                Require(factory.Disposed, "Native shutdown disposes session");
                Require(model.Connection.Map.LiveData is null, "Disconnect clears live positions");
                Capture(window, Path.Combine(output, "disconnected.png"));
                Console.WriteLine("PASS: native production shell, light/dark, 3 window sizes, terrain/live layers, tracking, chat, sidebar, recycler and disconnect.");
                result = 0;
            }
            catch (Exception error) { Console.Error.WriteLine(error); }
            finally
            {
                await model.ShutdownAsync();
                window.Close();
                stop.Cancel();
            }
        };
        window.Show();
        try { Dispatcher.UIThread.MainLoop(stop.Token); }
        catch (OperationCanceledException) when (stop.IsCancellationRequested) { }
        return result;
    }

    private static void Require(bool value, string label) { if (!value) throw new InvalidOperationException(label); }
    private static void Capture(Window window, string path)
    {
        using var image = new RenderTargetBitmap(new PixelSize((int)window.ClientSize.Width, (int)window.ClientSize.Height), new Vector(96, 96));
        image.Render(window);
        image.Save(path, PngBitmapEncoderOptions.Default);
    }

    private sealed class Catalog : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken)
            => Task.FromResult<IReadOnlyList<ServerSummary>>([new("Fixture Rust server", "fixture.example.test", 28082)
            { ProfileKey = "fixture", Devices = [new(1, "Compound lights", DeviceKind.Switch), new(2, "Base alarm", DeviceKind.Alarm)] }]);
    }
    private sealed class Picker : IProfileFilePicker { public Task<string?> PickAsync() => Task.FromResult<string?>(null); }
    private sealed class FixtureSessions : IServerSessionFactory, IServerSession
    {
        public bool Disposed { get; private set; }
        public IDeviceClient Devices { get; } = new TestDeviceClient { Read = (d, _) => Task.FromResult(new DeviceSnapshot(d.EntityId, false)) };
        public IServerChat Chat { get; } = new TestChatClient
        {
            Read = (channel, _) => Task.FromResult<IReadOnlyList<ChatMessage>>([new(channel, 1, "Fixture leader", "Fixture channel ready.", DateTimeOffset.UtcNow)]),
            Send = (channel, text, _) => Task.FromResult<ChatMessage?>(new(channel, 1, "Fixture leader", text, DateTimeOffset.UtcNow))
        };
        public IServerMap Map { get; } = new TestMapClient { Read = _ => Task.FromResult(Terrain()), ReadLive = _ => Task.FromResult(new MapLiveData(
            [new(10, MapMarkerKind.CargoShip, "Cargo ship", 600, 500), new(11, MapMarkerKind.PatrolHelicopter, "Patrol helicopter", 3000, 2500),
             new(12, MapMarkerKind.VendingMachine, "Fixture shop", 1100, 1300)],
            [new(1, "Fixture leader", 1800, 2000, true, true, true), new(2, "Fixture offline", 2700, 1500, false, true, false)])) };
        public ServerStatus InitialStatus { get; } = new("Fixture Rust server", 42, 200, 3, 14.5);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<IServerSession> ConnectAsync(string profilesPath, string profileKey, CancellationToken cancellationToken) => Task.FromResult<IServerSession>(this);
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public ValueTask DisposeAsync() { Disposed = true; return ValueTask.CompletedTask; }
    }

    private static ServerMapData Terrain()
    {
        using var bitmap = new WriteableBitmap(new PixelSize(512, 512), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Opaque);
        using (var buffer = bitmap.Lock())
        {
            var pixels = new byte[buffer.RowBytes * 512];
            for (int y = 0; y < 512; y++) for (int x = 0; x < 512; x++)
            {
                bool island = Math.Pow((x - 256) / 175.0, 2) + Math.Pow((y - 256) / 200.0, 2) < 1;
                int i = y * buffer.RowBytes + x * 4;
                pixels[i] = (byte)(island ? 65 : 120); pixels[i + 1] = (byte)(island ? 140 : 90);
                pixels[i + 2] = (byte)(island ? 90 : 50); pixels[i + 3] = 255;
            }
            Marshal.Copy(pixels, 0, buffer.Address, pixels.Length);
        }
        using var stream = new MemoryStream();
        bitmap.Save(stream, PngBitmapEncoderOptions.Default);
        return new(stream.ToArray(), 512, 512, 4000, [new("Fixture harbor", 1000, 2000), new("Fixture dome", 2800, 3200)]);
    }
}
