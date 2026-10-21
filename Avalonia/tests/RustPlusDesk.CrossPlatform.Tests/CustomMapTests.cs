using System.Text.Json.Nodes;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CustomMapTests
{
    [TestMethod]
    public void Url_IsTrimmed_EmptyClears_AndMustBeAbsolute()
    {
        Assert.IsTrue(CustomMapUrl.TryNormalize("   ", out string? url));
        Assert.IsNull(url);
        Assert.IsTrue(CustomMapUrl.TryNormalize("  https://maps.example.test/hd.png ", out url));
        Assert.AreEqual("https://maps.example.test/hd.png", url);
        Assert.IsFalse(CustomMapUrl.TryNormalize("maps.example.test/hd.png", out _));
        Assert.IsFalse(CustomMapUrl.TryNormalize("not a url", out _));
    }

    [TestMethod]
    public async Task Store_ChangesOnlyTheProfilesField_AndRejectsAMissingProfile()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyCustomMapUrlStore();
        Assert.IsNull(await store.GetAsync(fixture.Path, server.ProfileKey, default));
        var before = await fixture.JsonAsync();

        await store.SetAsync(fixture.Path, server.ProfileKey, "https://maps.example.test/hd.png", default);
        Assert.AreEqual("https://maps.example.test/hd.png", await store.GetAsync(fixture.Path, server.ProfileKey, default));
        var after = await fixture.JsonAsync();
        Assert.AreEqual("https://maps.example.test/hd.png", after[0]!["CustomMapUrl"]!.GetValue<string>());
        after[0]!.AsObject().Remove("CustomMapUrl");
        Assert.IsTrue(JsonNode.DeepEquals(before, after), "Every other field, unknown ones included, is kept.");

        await store.SetAsync(fixture.Path, server.ProfileKey, null, default);
        Assert.IsNull(await store.GetAsync(fixture.Path, server.ProfileKey, default));
        Assert.IsTrue((await fixture.JsonAsync())[0]!.AsObject().ContainsKey("CustomMapUrl"), "Cleared as null, like the reference serializer.");
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => store.SetAsync(fixture.Path, "missing-profile", "https://x.test/a.png", default));
    }

    private static byte[] Png(int size)
    {
        var target = new RenderTargetBitmap(new PixelSize(size, size));
        using (var context = target.CreateDrawingContext())
            context.FillRectangle(new SolidColorBrush(Color.FromRgb(90, 120, 60)), new Rect(0, 0, size, size));
        using var stream = new MemoryStream();
        target.Save(stream, PngBitmapEncoderOptions.Default);
        return stream.ToArray();
    }

    [TestMethod]
    public async Task Loader_UsesTheCachedOrDownloadedCustomImage_FallsBack_AndAWipeResetsTheUrl() => await UiTestSession.Current.Dispatch(async () =>
    {
        var cache = new MapCacheTests.MemoryCache();
        var downloader = new Downloader { Result = Png(512) };
        string? url = "https://maps.example.test/hd.png";
        int resets = 0;
        var log = new List<string>();
        var loader = new MapCacheLoader(cache, null, () => ("10.0.0.1", 28082))
        {
            Log = log.Add,
            Downloader = downloader,
            CustomMapUrl = _ => Task.FromResult<string?>(url),
            ResetCustomMapUrl = _ => { resets++; url = null; return Task.CompletedTask; }
        };
        var wipe = new DateTimeOffset(2026, 10, 1, 0, 0, 0, TimeSpan.Zero);
        var server = MapTests.SyntheticMap();

        var map = await loader.LoadAsync(new MapCacheTests.Session(server, wipe), default);
        Assert.IsTrue(map.CustomImage);
        Assert.AreEqual((512, 512, 4500u), (map.Width, map.Height, map.WorldSize));
        CollectionAssert.AreEqual(server.Monuments.ToArray(), map.Monuments.ToArray(), "Monuments stay the server's.");
        Assert.AreEqual(1024, cache.Entries["10.0.0.1:28082"].PixelWidth, "The server map itself is cached, not the custom image.");
        Assert.IsTrue(cache.Custom.ContainsKey("10.0.0.1:28082"));
        Assert.AreEqual("https://maps.example.test/hd.png", downloader.Urls.Single());
        Assert.IsTrue(log.Contains("[custom-map] Downloading custom HD map image from URL: https://maps.example.test/hd.png"));

        var model = new MapViewModel();
        model.Apply(map);
        Assert.AreEqual(MapGeometry.CustomMapPadding, model.Geometry!.PadWorld, "Custom images use the reference 1000 m margin.");
        model.Apply(server);
        Assert.AreEqual(MapGeometry.ServerMapPadding, model.Geometry!.PadWorld);

        Assert.IsTrue((await loader.LoadAsync(new MapCacheTests.Session(server, wipe), default)).CustomImage);
        Assert.AreEqual(1, downloader.Urls.Count, "The cached custom image is reused.");

        cache.DeleteCustom("10.0.0.1", 28082);
        downloader.Result = [1, 2, 3];
        Assert.IsFalse((await loader.LoadAsync(new MapCacheTests.Session(server, wipe), default)).CustomImage, "Not an image: the server map stays.");
        Assert.IsTrue(log.Contains("[custom-map] Failed to download image from URL. Falling back to server map."));
        Assert.IsFalse(cache.Custom.ContainsKey("10.0.0.1:28082"));

        downloader.Result = Png(256);
        var afterWipe = await loader.LoadAsync(new MapCacheTests.Session(server, wipe.AddDays(7)), default);
        Assert.IsFalse(afterWipe.CustomImage, "A wipe drops the custom map.");
        Assert.AreEqual((1, (string?)null), (resets, url));
        Assert.IsTrue(log.Contains("[map-cache] Wipe detected; resetting Custom HD Map URL for current server profile."));
        return true;
    }, default);

    [TestMethod]
    public async Task SettingsCard_ValidatesSavesClearsAndReloads() => await UiTestSession.Current.Dispatch(async () =>
    {
        var store = new MemoryUrls();
        var cache = new MapCacheTests.MemoryCache();
        cache.SaveCustom("10.0.0.1", 28082, [1]);
        ServerSummary? selected = null;
        int reloads = 0;
        var card = new CustomMapSettingsViewModel(store, cache, () => selected is null ? null : ("profiles.json", selected), () => { reloads++; return Task.CompletedTask; });
        await card.LoadAsync();
        Assert.AreEqual(CustomMapUrl.NoServerStatus, card.Status);
        await card.ApplyCommand.ExecuteAsync(null);
        Assert.IsTrue(card.IsError);

        selected = new ServerSummary("Synthetic", "10.0.0.1", 28082) { ProfileKey = "key" };
        await card.LoadAsync();
        Assert.AreEqual((CustomMapUrl.NoneStatus, ""), (card.Status, card.Url));
        card.Url = "not a url";
        await card.ApplyCommand.ExecuteAsync(null);
        Assert.AreEqual((CustomMapUrl.InvalidMessage, true, 0), (card.Status, card.IsError, store.Values.Count));

        card.Url = " https://maps.example.test/hd.png ";
        await card.ApplyCommand.ExecuteAsync(null);
        Assert.AreEqual("https://maps.example.test/hd.png", store.Values["key"]);
        Assert.AreEqual((CustomMapUrl.UpdatedStatus, false, 1), (card.Status, card.IsError, reloads));
        Assert.IsFalse(cache.Custom.ContainsKey("10.0.0.1:28082"), "A new URL downloads its image again.");
        await card.LoadAsync();
        Assert.AreEqual((CustomMapUrl.ActiveStatus, "https://maps.example.test/hd.png"), (card.Status, card.Url));

        var settings = new SettingsViewModel(new ShopNoticeTests.MemorySettings()) { CustomMap = card };
        var window = new Window { Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark,
            Content = new Border { Padding = new Thickness(16), Child = new SettingsView { DataContext = settings } } };
        try
        {
            window.Show();
            window.UpdateLayout();
            var titles = window.GetVisualDescendants().OfType<TextBlock>().Where(t => t.Classes.Contains("title") && t.IsEffectivelyVisible).Select(t => t.Text).ToList();
            CollectionAssert.AreEqual(new[] { "General", "Map", "Alerts", "Maintenance", "Credits" }, titles);
            var box = window.GetVisualDescendants().OfType<TextBox>().Single(t => AutomationProperties.GetName(t) == "Custom map image URL");
            Assert.AreEqual("https://maps.example.test/hd.png", box.Text);
            var clear = window.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "Clear custom map URL");
            clear.BringIntoView();
            window.UpdateLayout();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-settings-map.png"), PngBitmapEncoderOptions.Default);
            }
            await card.ClearCommand.ExecuteAsync(null);
            Assert.IsNull(store.Values["key"]);
            Assert.AreEqual((CustomMapUrl.ClearedStatus, "", 2), (card.Status, box.Text, reloads));
        }
        finally { window.Close(); }
        return true;
    }, default);

    private sealed class Downloader : ICustomMapDownloader
    {
        public byte[]? Result { get; set; }
        public List<string> Urls { get; } = [];
        public Task<byte[]?> DownloadAsync(string url, CancellationToken ct) { Urls.Add(url); return Task.FromResult(Result); }
    }

    private sealed class MemoryUrls : ICustomMapUrlStore
    {
        public Dictionary<string, string?> Values { get; } = [];
        public Task<string?> GetAsync(string profilesPath, string profileKey, CancellationToken ct) => Task.FromResult(Values.GetValueOrDefault(profileKey));
        public Task SetAsync(string profilesPath, string profileKey, string? url, CancellationToken ct) { Values[profileKey] = url; return Task.CompletedTask; }
    }
}
