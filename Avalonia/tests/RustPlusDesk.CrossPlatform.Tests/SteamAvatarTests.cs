using System.Net;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Avalonia.Shared;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Infrastructure.Steam;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class SteamAvatarTests
{
    internal static byte[] Png(Color color, int size = 64)
    {
        var target = new RenderTargetBitmap(new PixelSize(size, size));
        using (var context = target.CreateDrawingContext())
        {
            context.FillRectangle(new SolidColorBrush(color), new Rect(0, 0, size, size));
            context.DrawEllipse(Brushes.White, null, new Point(size / 2.0, size / 2.6), size / 5.0, size / 5.0);
        }
        using var stream = new MemoryStream();
        target.Save(stream, PngBitmapEncoderOptions.Default);
        return stream.ToArray();
    }

    [TestMethod]
    public void Xml_PrefersTheFullAvatar()
    {
        Assert.AreEqual("https://avatars.example.test/full.jpg",
            SteamAvatarXml.AvatarUrl("<profile><avatarMedium><![CDATA[https://avatars.example.test/medium.jpg]]></avatarMedium><AvatarFull><![CDATA[https://avatars.example.test/full.jpg]]></AvatarFull></profile>"));
        Assert.AreEqual("https://avatars.example.test/medium.jpg", SteamAvatarXml.AvatarUrl("<avatarMedium><![CDATA[https://avatars.example.test/medium.jpg]]></avatarMedium>"));
        Assert.IsNull(SteamAvatarXml.AvatarUrl("<response><error>The specified profile could not be found.</error></response>"));
    }

    [TestMethod]
    public async Task Source_LoadsTheProfileXmlThenTheImage_AndFailsQuietly()
    {
        var requests = new List<string>();
        string xml = "<avatarFull><![CDATA[https://avatars.example.test/a.jpg]]></avatarFull>";
        var status = HttpStatusCode.OK;
        using var http = new HttpClient(new Handler(request =>
        {
            requests.Add(request.RequestUri!.ToString());
            return new HttpResponseMessage(status)
            {
                Content = request.RequestUri.Host == "steamcommunity.com" ? new StringContent(xml) : new ByteArrayContent([1, 2, 3])
            };
        }));
        var source = new SteamCommunityAvatarSource(http);
        CollectionAssert.AreEqual(new byte[] { 1, 2, 3 }, await source.LoadAsync(76561198000000001, default));
        CollectionAssert.AreEqual(new[] { "https://steamcommunity.com/profiles/76561198000000001?xml=1", "https://avatars.example.test/a.jpg" }, requests);
        Assert.IsNull(await source.LoadAsync(0, default));
        xml = "<avatarFull><![CDATA[http://avatars.example.test/a.jpg]]></avatarFull>";
        Assert.IsNull(await source.LoadAsync(76561198000000001, default), "Only https avatar URLs are fetched.");
        xml = "<profile />";
        Assert.IsNull(await source.LoadAsync(76561198000000001, default));
        status = HttpStatusCode.TooManyRequests;
        Assert.IsNull(await source.LoadAsync(76561198000000001, default));
    }

    [TestMethod]
    public async Task Cache_DeduplicatesAndRetriesAfterThirtySeconds() => await UiTestSession.Current.Dispatch(async () =>
    {
        var now = new DateTime(2026, 10, 18, 12, 0, 0, DateTimeKind.Utc);
        var source = new FakeSource();
        var cache = new AvatarCache(source, () => now);
        var loaded = new List<ulong>();
        cache.Loaded += loaded.Add;
        var gate = new TaskCompletionSource<byte[]?>();
        source.Next = () => gate.Task;
        var first = cache.RequestAsync(7);
        await cache.RequestAsync(7);
        Assert.AreEqual(1, source.Calls, "One request per Steam ID at a time.");
        gate.SetResult(null);
        await first;
        Assert.IsNull(cache.Get(7));
        await cache.RequestAsync(7);
        Assert.AreEqual(1, source.Calls, "A missing avatar waits 30 seconds.");
        now = now.AddSeconds(30);
        source.Next = () => Task.FromResult<byte[]?>([1, 2, 3]);
        await cache.RequestAsync(7);
        Assert.AreEqual((2, (Bitmap?)null), (source.Calls, cache.Get(7)), "Bytes that are not an image are a failure too.");
        now = now.AddSeconds(31);
        source.Next = () => Task.FromResult<byte[]?>(Png(Colors.SteelBlue));
        await cache.RequestAsync(7);
        Assert.IsNotNull(cache.Get(7));
        CollectionAssert.AreEqual(new ulong[] { 7 }, loaded);
        Assert.AreEqual(1, cache.Version);
        await cache.RequestAsync(7);
        Assert.AreEqual(3, source.Calls, "A cached avatar is not loaded again.");
        return true;
    }, default);

    [TestMethod]
    public async Task TeamDockAndMap_ShowLoadedAvatars() => await UiTestSession.Current.Dispatch(async () =>
    {
        var avatar = new Bitmap(new MemoryStream(Png(Colors.SteelBlue)));
        var team = new TeamViewModel
        {
            IsExpanded = true,
            Message = "Current server team",
            Members = [new TeamMemberRow(new ServerTeamMember("76561198000000001", "Synthetic Builder", true, true)),
                new TeamMemberRow(new ServerTeamMember("76561198000000002", "Farmer", false, true))]
        };
        Assert.IsFalse(team.Members[0].HasAvatar);
        Assert.AreEqual("S", team.Members[0].Initial);
        team.Avatars = id => id == "76561198000000001" ? avatar : null;
        team.RefreshAvatars();
        Assert.IsTrue(team.Members[0].HasAvatar);
        Assert.IsFalse(team.Members[1].HasAvatar);
        var dock = new Window { Width = 360, Height = 220, RequestedThemeVariant = ThemeVariant.Dark, Content = new TeamView { DataContext = team } };
        try
        {
            dock.Show();
            Assert.AreEqual(1, dock.GetVisualDescendants().OfType<Image>().Count(image => image.IsEffectivelyVisible && image.Source is not null));
            using var frame = dock.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-team-avatars.png"), PngBitmapEncoderOptions.Default);
        }
        finally { dock.Close(); }

        var map = new MapViewModel();
        map.Apply(MapTests.SyntheticMap());
        var surface = new MapSurface
        {
            Image = map.Image, Geometry = map.Geometry, Avatars = id => id == 76561198000000001 ? avatar : null,
            Markers = [new(1, MapMarkerKind.Player, 2200, 2300) { SteamId = 76561198000000001, Name = "Synthetic Builder" },
                new(2, MapMarkerKind.Player, 3200, 1500) { SteamId = 76561198000000002, Name = "Farmer" }]
        };
        var window = new Window { Width = 700, Height = 600, Content = surface };
        try
        {
            window.Show();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-avatars.png"), PngBitmapEncoderOptions.Default);
            }
            surface.ShowSteamAvatars = false;
            using var plain = window.CaptureRenderedFrame();
            Assert.IsNotNull(plain);
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class FakeSource : ISteamAvatarSource
    {
        public Func<Task<byte[]?>> Next { get; set; } = () => Task.FromResult<byte[]?>(null);
        public int Calls { get; private set; }
        public Task<byte[]?> LoadAsync(ulong steamId, CancellationToken ct) { Calls++; return Next(); }
    }

    private sealed class Handler(Func<HttpRequestMessage, HttpResponseMessage> send) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => Task.FromResult(send(request));
    }
}
