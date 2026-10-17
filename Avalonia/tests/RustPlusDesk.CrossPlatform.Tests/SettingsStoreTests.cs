using System.Text.Json.Nodes;
using Avalonia.Controls;
using Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Settings;
using RustPlusDesk.Infrastructure.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class SettingsStoreTests
{
    [TestMethod]
    public async Task Store_ReadsReferenceKeys_WritesOnlyChanges_AndKeepsUnknownAndExternalKeys()
    {
        using var file = new SettingsFile("""{"MapGridOpacity":0.4,"SteamId64":"76561198000000001","FutureSetting":{"a":[1,2]}}""");
        var store = new TrackingSettingsStore(file.Path, TimeSpan.FromMilliseconds(10));
        Assert.AreEqual(0.4, store.Get(SettingKeys.MapGridOpacity, 0.7));
        Assert.AreEqual(1.0, store.Get(SettingKeys.MapMonumentScale, 1.0));
        Assert.AreEqual(3, store.Get("SteamId64", 3), "A wrongly typed value falls back instead of throwing.");
        // The WPF app changes another key between our read and our write.
        var external = JsonNode.Parse(await File.ReadAllTextAsync(file.Path))!;
        external["AnnounceCargo"] = true;
        await File.WriteAllTextAsync(file.Path, external.ToJsonString());
        store.Set(SettingKeys.MapMonumentScale, 2.5);
        await store.FlushAsync();
        var saved = JsonNode.Parse(await File.ReadAllTextAsync(file.Path))!;
        Assert.AreEqual(2.5, saved["MapMonumentScale"]!.GetValue<double>());
        Assert.AreEqual(0.4, saved["MapGridOpacity"]!.GetValue<double>());
        Assert.IsTrue(saved["AnnounceCargo"]!.GetValue<bool>());
        Assert.IsTrue(JsonNode.DeepEquals(JsonNode.Parse("""{"a":[1,2]}"""), saved["FutureSetting"]));
        Assert.IsTrue(store.Get("AnnounceCargo", false), "External changes are adopted after a save.");
        Assert.AreEqual(1, Directory.GetFiles(file.Directory).Length);
    }

    [TestMethod]
    public async Task Store_DebouncesWrites_SkipsUnchangedValues_AndBacksUpUnreadableFiles()
    {
        using var file = new SettingsFile("{ not json");
        var store = new TrackingSettingsStore(file.Path, TimeSpan.FromMilliseconds(20));
        Assert.AreEqual(0.7, store.Get(SettingKeys.MapGridOpacity, 0.7));
        store.Set(SettingKeys.RailShowLabels, true);
        store.Set(SettingKeys.RailShowLabels, true);
        await CompanionListenerViewModelTests.Eventually(() => File.ReadAllText(file.Path).Contains("RailShowLabels"));
        Assert.AreEqual("{ not json", await File.ReadAllTextAsync(Directory.GetFiles(file.Directory, "*.bak").Single()));
        Assert.IsTrue(JsonNode.Parse(await File.ReadAllTextAsync(file.Path))!["RailShowLabels"]!.GetValue<bool>());
    }

    [TestMethod]
    public async Task MapLayers_LoadClampedReferenceDefaults_AndWriteChangesBack()
    {
        using var file = new SettingsFile("""{"MapMonumentScale":9,"MapGridOpacity":0.05,"MapMonumentDisplayMode":1,"AutoLoadShops":false}""");
        var store = new TrackingSettingsStore(file.Path, TimeSpan.FromMilliseconds(10));
        var map = new MapViewModel();
        map.UseSettings(store);
        Assert.AreEqual(3, map.MonumentScale);
        Assert.AreEqual(0.1, map.GridOpacity);
        Assert.IsTrue(map.ShowMonumentNames);
        Assert.IsFalse(map.ShowShops, "AutoLoadShops off starts with shops hidden.");
        Assert.AreEqual("""{"MapMonumentScale":9,"MapGridOpacity":0.05,"MapMonumentDisplayMode":1,"AutoLoadShops":false}""",
            await File.ReadAllTextAsync(file.Path), "Loading alone never rewrites the file.");
        map.MonumentDisplayMode = 0;
        map.MonumentOpacity = 0.5;
        await store.FlushAsync();
        var saved = JsonNode.Parse(await File.ReadAllTextAsync(file.Path))!;
        Assert.AreEqual(0, saved["MapMonumentDisplayMode"]!.GetValue<int>());
        Assert.AreEqual(0.5, saved["MapMonumentOpacity"]!.GetValue<double>());
    }

    [TestMethod]
    public async Task MapSurface_RendersMonumentNamesStyleAndOpacities()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new MapViewModel();
            model.Apply(MapTests.SyntheticMap());
            var surface = new MapSurface
            {
                Image = model.Image, Geometry = model.Geometry, Monuments = model.Monuments,
                MonumentNames = true, MonumentScale = 2, MonumentOpacity = 0.6, GridOpacity = 0.3
            };
            var window = new Window { Width = 900, Height = 700, Content = surface };
            try
            {
                window.Show();
                using var frame = window.CaptureRenderedFrame();
                Assert.IsNotNull(frame);
                frame!.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-names.png"), global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }

    private sealed class SettingsFile : IDisposable
    {
        public string Directory { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rustplus-settings-" + Guid.NewGuid().ToString("N"));
        public string Path => System.IO.Path.Combine(Directory, "tracking_settings.json");
        public SettingsFile(string contents) { System.IO.Directory.CreateDirectory(Directory); File.WriteAllText(Path, contents); }
        public void Dispose() => System.IO.Directory.Delete(Directory, recursive: true);
    }
}
