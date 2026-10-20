using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Avalonia.Features.Hotkeys;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Hotkeys;
using RustPlusDesk.Infrastructure.Hotkeys;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class HotkeyTests
{
    private const string Server = "10.0.0.1:28082";

    private sealed class NoPicker : IProfileFilePicker
    {
        public Task<string?> PickAsync() => throw new AssertFailedException("The file picker should not run in this test.");
    }

    private sealed class FakeHotkeys(bool supported = true) : IGlobalHotkeys
    {
        public HashSet<string> Taken { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> Registered { get; } = [];
        private readonly Dictionary<string, bool> status = new(StringComparer.OrdinalIgnoreCase);
        public bool IsSupported => supported;
        public string UnavailableReason => supported ? "" : "Not on this system.";
        public bool Register(string gesture)
        {
            bool ok = supported && !Taken.Contains(gesture);
            status[gesture] = ok;
            if (ok) Registered.Add(gesture);
            return ok;
        }
        public void UnregisterAll() { Registered.Clear(); status.Clear(); }
        public IReadOnlyDictionary<string, bool> RegistrationStatus => status;
        public event Action<string>? Pressed;
        public void Press(string gesture) => Pressed?.Invoke(gesture);
    }

    private sealed class MemoryStore : IHotkeyStore
    {
        public HotkeyMap Map { get; set; } = new();
        public HotkeyOptions Options { get; set; } = new();
        public int Saves { get; private set; }
        public Task<HotkeyMap> LoadMapAsync(CancellationToken ct) => Task.FromResult(Map.Clone());
        public Task SaveMapAsync(HotkeyMap map, CancellationToken ct) { Map = map.Clone(); Saves++; return Task.CompletedTask; }
        public Task<HotkeyOptions> LoadOptionsAsync(CancellationToken ct) => Task.FromResult(Options);
        public Task SaveOptionsAsync(HotkeyOptions options, CancellationToken ct) { Options = options; return Task.CompletedTask; }
    }

    [TestMethod]
    public void Map_MovesDevicesBetweenGestures_AndPrunesEmptyOnes()
    {
        var map = new HotkeyMap();
        Assert.AreEqual("default", HotkeyMap.ServerKey(null, null));
        Assert.AreEqual(Server, HotkeyMap.ServerKey("10.0.0.1", 28082));
        map.Assign(Server, "Ctrl+Alt+K", 1);
        map.Assign(Server, "Ctrl+Alt+K", 2);
        map.Assign(Server, "F5", 1);
        Assert.AreEqual("F5", map.GestureFor(Server, 1), "A device belongs to one gesture per server.");
        CollectionAssert.AreEqual(new long[] { 2 }, map.ForServer(Server)["Ctrl+Alt+K"]);
        map.Clear(Server, 2);
        Assert.IsNull(map.GestureFor(Server, 2));
        var copy = map.Clone();
        map.Prune();
        Assert.IsFalse(map.ForServer(Server).ContainsKey("Ctrl+Alt+K"));
        Assert.IsTrue(copy.ForServer(Server).ContainsKey("Ctrl+Alt+K"), "Clones are independent.");
        Assert.IsTrue(map.Remove(Server, [1]));
        Assert.AreEqual(0, map.ForServer(Server).Count);
        Assert.IsFalse(map.Remove("other", [1]));

        Assert.AreEqual("Ctrl+Alt+Shift+Win+F5", HotkeyGesture.Format("F5", true, true, true, true));
        Assert.AreEqual("NumPad1", HotkeyGesture.Format("NumPad1", false, false, false, false));
        Assert.IsTrue(HotkeyGesture.TryParse("ctrl + meta + D1", out bool ctrl, out bool alt, out bool shift, out bool win, out string key));
        Assert.IsTrue(ctrl && win && !alt && !shift);
        Assert.AreEqual("D1", key);
        Assert.IsFalse(HotkeyGesture.TryParse("Ctrl+Alt", out _, out _, out _, out _, out _));
    }

    [TestMethod]
    public void Win32_ParsesGesturesToVirtualKeys()
    {
        Assert.IsTrue(Win32GlobalHotkeys.TryParse("Ctrl+Alt+K", out uint modifiers, out uint vk));
        Assert.AreEqual(0x0003u, modifiers);
        Assert.AreEqual(0x4Bu, vk);
        Assert.IsTrue(Win32GlobalHotkeys.TryParse("Shift+Win+NumPad7", out modifiers, out vk));
        Assert.AreEqual(0x000Cu, modifiers);
        Assert.AreEqual(0x67u, vk);
        Assert.IsTrue(Win32GlobalHotkeys.TryParse("F24", out _, out vk));
        Assert.AreEqual(0x87u, vk);
        Assert.IsTrue(Win32GlobalHotkeys.TryParse("D0", out _, out vk));
        Assert.AreEqual(0x30u, vk);
        Assert.IsTrue(Win32GlobalHotkeys.TryParse("OemPlus", out _, out vk));
        Assert.AreEqual(0xBBu, vk);
        Assert.IsFalse(Win32GlobalHotkeys.TryParse("Ctrl+NotAKey", out _, out _));
        Assert.IsFalse(Win32GlobalHotkeys.TryParse("Ctrl+VolumeUp", out _, out _), "Keys without a mapping are refused, not guessed.");
    }

    [TestMethod]
    public async Task Store_SharesTheReferenceFiles_AndMigratesTheFlatShape()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-hotkeys-" + Guid.NewGuid().ToString("N"));
        try
        {
            var store = new HotkeyStore(folder);
            Assert.AreEqual(0, (await store.LoadMapAsync(CancellationToken.None)).Servers.Count);
            Assert.AreEqual(150, (await store.LoadOptionsAsync(CancellationToken.None)).ToggleDelayMs);
            var map = new HotkeyMap();
            map.Assign(Server, "Ctrl+Alt+K", 12);
            map.ForServer(Server)["Unused"] = [];
            await store.SaveMapAsync(map, CancellationToken.None);
            string json = await File.ReadAllTextAsync(Path.Combine(folder, "hotkeys.json"));
            StringAssert.Contains(json, "\"10.0.0.1:28082\"");
            Assert.IsFalse(json.Contains("Unused"), "Empty gestures are pruned on save.");
            Assert.IsTrue(map.ForServer(Server).ContainsKey("Unused"), "Saving never changes the caller's map.");
            Assert.AreEqual("Ctrl+Alt+K", (await store.LoadMapAsync(CancellationToken.None)).GestureFor(Server, 12));

            await store.SaveOptionsAsync(new HotkeyOptions { ParallelMode = true, ToggleDelayMs = 300 }, CancellationToken.None);
            var options = await store.LoadOptionsAsync(CancellationToken.None);
            Assert.IsTrue(options.ParallelMode);
            Assert.AreEqual(300, options.ToggleDelayMs);
            StringAssert.Contains(await File.ReadAllTextAsync(Path.Combine(folder, "hotkey_options.json")), "\"ToggleDelayMs\": 300");

            await File.WriteAllTextAsync(Path.Combine(folder, "hotkeys.json"), "{ \"Ctrl+F1\": [5, 6] }");
            var legacy = await store.LoadMapAsync(CancellationToken.None);
            CollectionAssert.AreEqual(new long[] { 5, 6 }, legacy.ForServer(HotkeyMap.DefaultServer)["Ctrl+F1"]);
            await File.WriteAllTextAsync(Path.Combine(folder, "hotkeys.json"), "not json");
            Assert.AreEqual(0, (await store.LoadMapAsync(CancellationToken.None)).Servers.Count);
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public async Task Controller_ActivatesPerServer_DebouncesPresses_AndRunsOneSequenceAtATime()
    {
        var store = new MemoryStore();
        store.Map.Assign(Server, "Ctrl+Alt+K", 1);
        store.Map.Assign(Server, "Ctrl+Alt+K", 2);
        store.Map.Assign("other:1", "F9", 9);
        var now = DateTimeOffset.UnixEpoch;
        var model = new HotkeysViewModel(store, () => now);
        var fake = new FakeHotkeys();
        model.UseGlobalHotkeys(fake);
        model.SelectServer("10.0.0.1", 28082);
        var gate = new TaskCompletionSource<int>();
        var toggles = new List<uint[]>();
        model.Toggle = (ids, parallel, delay) => { toggles.Add(ids.ToArray()); Assert.AreEqual(150, delay); return gate.Task; };
        var log = new List<string>();
        model.Log = log.Add;
        await model.LoadAsync(CancellationToken.None);
        Assert.IsTrue(model.IsActive);
        Assert.AreEqual("Hotkeys active", model.ButtonText);
        CollectionAssert.AreEqual(new[] { "Ctrl+Alt+K" }, fake.Registered, "Only the selected server's gestures register.");

        fake.Press("Ctrl+Alt+K");
        CollectionAssert.AreEqual(new uint[] { 1, 2 }, toggles.Single());
        now = now.AddMilliseconds(100);
        fake.Press("Ctrl+Alt+K");
        Assert.AreEqual(1, toggles.Count, "Presses within 350 ms are debounced.");
        now = now.AddMilliseconds(400);
        fake.Press("Ctrl+Alt+K");
        Assert.AreEqual(1, toggles.Count, "A running sequence ignores further presses.");
        CollectionAssert.Contains(log, "[hotkeys] Hotkey sequence already running – ignored.");
        gate.SetResult(2);
        await model.WaitIdleAsync();
        now = now.AddMilliseconds(400);
        fake.Press("Ctrl+Alt+K");
        Assert.AreEqual(2, toggles.Count);
        fake.Press("F9");
        Assert.AreEqual(2, toggles.Count, "Gestures of other servers do nothing.");

        model.SelectServer("other", 1);
        CollectionAssert.AreEqual(new[] { "F9" }, fake.Registered, "Switching servers re-registers while active.");
        fake.Taken.Add("F9");
        model.Activate();
        Assert.IsFalse(model.IsActive, "Nothing registered: not active.");
        CollectionAssert.Contains(log, "[hotkeys] Cannot register hotkey 'F9'.");
        model.Deactivate();
        Assert.AreEqual("Hotkeys", model.ButtonText);
    }

    [TestMethod]
    public async Task Editor_ListsSwitchesAndDeletedIds_AppliesOnClose()
    {
        var store = new MemoryStore();
        store.Map.Assign(Server, "Ctrl+Alt+K", 1);
        store.Map.Assign(Server, "Ctrl+Alt+K", 77);
        var model = new HotkeysViewModel(store);
        var fake = new FakeHotkeys();
        fake.Taken.Add("Ctrl+Alt+K");
        model.UseGlobalHotkeys(fake);
        model.SelectServer("10.0.0.1", 28082);
        await model.LoadAsync(CancellationToken.None);
        var devices = new DevicesViewModel();
        devices.Select([new(0, "Base", DeviceKind.Group) { Children = [new(1, "Lights", DeviceKind.Switch), new(2, "Pump", DeviceKind.Switch), new(3, "Door", DeviceKind.Alarm)] }]);
        model.Devices = () => devices.Devices.SelectMany(row => row.DescendantsAndSelf());
        HotkeysEditorViewModel? editor = null;
        model.EditRequested += (_, value) => editor = value;
        model.EditCommand.Execute(null);
        Assert.IsNotNull(editor);
        Assert.IsFalse(model.IsActive, "Editing deactivates hotkeys first.");
        CollectionAssert.AreEqual(new[] { "Lights", "Pump", "Unknown Device" }, editor.Rows.Select(row => row.DisplayName).ToArray());
        var lights = editor.Rows[0];
        Assert.AreEqual("Ctrl+Alt+K", lights.Hotkey);
        Assert.IsTrue(lights.IsRegisteredWarningVisible, "Another app held the key at the last activation.");
        Assert.AreEqual("–", lights.Status);
        Assert.AreEqual("ID: 1", lights.IdText);
        var deleted = editor.Rows[2];
        Assert.IsTrue(deleted.IsDeleted);
        Assert.AreEqual("DELETED", deleted.Status);
        Assert.AreEqual(77, deleted.EntityId);

        editor.Assign(editor.Rows[1], "F6");
        editor.ClearCommand.Execute(deleted);
        editor.IsParallel = true;
        Assert.IsFalse(editor.IsSequential);
        editor.DelayMs = 333;
        Assert.AreEqual(350, editor.DelayMs, "The delay snaps to 50 ms steps.");
        editor.DelayMs = 5000;
        Assert.AreEqual(1000, editor.DelayMs);
        Assert.AreEqual("Toggle delay: 1000 ms", editor.DelayText);
        Assert.AreEqual("Ctrl+Alt+K", store.Map.GestureFor(Server, 1), "Nothing is written until the window closes.");
        editor.CloseActivateCommand.Execute(null);
        model.Apply(editor);
        await model.FlushAsync();
        Assert.AreEqual("F6", store.Map.GestureFor(Server, 2));
        Assert.IsNull(store.Map.GestureFor(Server, 77));
        Assert.IsTrue(store.Options.ParallelMode);
        Assert.AreEqual(1000, store.Options.ToggleDelayMs);
        Assert.IsTrue(model.IsActive, "F6 registered.");

        model.RemoveDevices("10.0.0.1", 28082, [2]);
        await model.FlushAsync();
        Assert.IsNull(store.Map.GestureFor(Server, 2), "Deleted devices lose their hotkeys.");

        var unsupported = new HotkeysViewModel(new MemoryStore());
        unsupported.UseGlobalHotkeys(new UnavailableGlobalHotkeys("Not on this system."));
        HotkeysEditorViewModel? limited = null;
        unsupported.EditRequested += (_, value) => limited = value;
        unsupported.EditCommand.Execute(null);
        Assert.IsFalse(limited!.CanActivate);
        Assert.IsFalse(limited.CloseActivateCommand.CanExecute(null), "Activation is not offered where it cannot work.");
        Assert.AreEqual("Not on this system.", limited.UnavailableReason);
        StringAssert.Contains(UnavailableGlobalHotkeys.ForCurrentPlatform().UnavailableReason, "Assignments are saved");
    }

    [TestMethod]
    public async Task Devices_ToggleFlipsUniqueSwitches_WithTheSequentialDelay()
    {
        var calls = new List<(uint Id, bool Value)>();
        var client = new TestDeviceClient
        {
            Read = (device, _) => Task.FromResult(new DeviceSnapshot(device.EntityId, device.EntityId == 1)),
            Set = (device, value, _) => { calls.Add((device.EntityId, value)); return Task.FromResult(new DeviceSnapshot(device.EntityId, value)); }
        };
        var model = new DevicesViewModel();
        model.Select([new(0, "Base", DeviceKind.Group) { Children = [new(1, "Lights", DeviceKind.Switch), new(2, "Pump", DeviceKind.Switch), new(3, "Door", DeviceKind.Alarm)] },
            new(2, "Pump alias", DeviceKind.Switch)]);
        Assert.AreEqual(0, await model.ToggleAsync([1], false, 0), "Nothing happens before connecting.");
        using var lifetime = new CancellationTokenSource();
        var running = model.RunAsync(client, lifetime.Token);
        try
        {
            await CompanionListenerViewModelTests.Eventually(() => model.Devices[0].Children[0].IsOn == true);
            var started = DateTime.UtcNow;
            int confirmed = await model.ToggleAsync([0, 2, 99], false, 60);
            Assert.AreEqual(2, confirmed);
            CollectionAssert.AreEqual(new[] { (1u, false), (2u, true) }, calls, "The group expands to its switches; each flips once.");
            Assert.IsTrue(DateTime.UtcNow - started >= TimeSpan.FromMilliseconds(50), "Sequential mode waits between switches.");
            Assert.AreEqual(true, model.Devices[1].IsOn, "Aliases follow the confirmed state.");
            calls.Clear();
            await model.ToggleAsync([2], true, 1000);
            CollectionAssert.AreEqual(new[] { (2u, false) }, calls);
        }
        finally { lifetime.Cancel(); await running; }
    }

    [TestMethod]
    public async Task Windows_RenderTheEditor_AndCaptureGestures()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new HotkeyMap();
            map.Assign(Server, "Ctrl+Alt+K", 1);
            map.Assign(Server, "Ctrl+Alt+K", 77);
            var devices = new DevicesViewModel();
            devices.Select([new(1, "Lights", DeviceKind.Switch), new(2, "Pump", DeviceKind.Switch)]);
            var editor = new HotkeysEditorViewModel(map, Server, new HotkeyOptions(), devices.Devices, new Dictionary<string, bool> { ["Ctrl+Alt+K"] = false },
                false, UnavailableGlobalHotkeys.ForCurrentPlatform().UnavailableReason);
            var window = new HotkeysWindow { DataContext = editor, Width = 640, Height = 560 };
            var capture = new HotkeyCaptureWindow();
            try
            {
                window.Show();
                window.UpdateLayout();
                Assert.AreEqual(3, window.GetVisualDescendants().OfType<Button>().Count(button => AutomationProperties.GetName(button) == "Set hotkey"));
                var activate = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Close and activate hotkeys");
                Assert.IsFalse(activate.IsEffectivelyEnabled);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-hotkeys.png"), PngBitmapEncoderOptions.Default);
                }
                bool closed = false;
                window.Closed += (_, _) => closed = true;
                editor.CloseDeactivateCommand.Execute(null);
                Assert.IsTrue(closed, "The editor closes itself.");
                Assert.AreEqual(false, editor.Result);

                capture.Show();
                capture.OnKeyDown(capture, new KeyEventArgs { Key = Key.LeftCtrl, KeyModifiers = KeyModifiers.Control, RoutedEvent = InputElement.KeyDownEvent });
                Assert.IsNull(capture.Gesture, "Modifiers alone wait for a key.");
                capture.OnKeyDown(capture, new KeyEventArgs { Key = Key.Enter, RoutedEvent = InputElement.KeyDownEvent });
                Assert.IsTrue(capture.IsVisible, "Enter needs a combination first.");
                capture.OnKeyDown(capture, new KeyEventArgs { Key = Key.K, KeyModifiers = KeyModifiers.Control | KeyModifiers.Alt, RoutedEvent = InputElement.KeyDownEvent });
                Assert.AreEqual("Ctrl+Alt+K", capture.Gesture);
                capture.OnKeyDown(capture, new KeyEventArgs { Key = Key.NumPad3, KeyModifiers = KeyModifiers.Meta, RoutedEvent = InputElement.KeyDownEvent });
                Assert.AreEqual("Win+NumPad3", capture.Gesture);
            }
            finally { window.Close(); capture.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }

    /// <summary>
    /// The devices toolbar is a row of icon buttons, and the hotkeys are one of them: the button that opened the
    /// editor used to be a bordered pill in the middle of the row, which broke the toolbar's shape.
    /// </summary>
    [TestMethod]
    public async Task TheDevicesToolbar_KeepsOneShape_AndTheHotkeysButtonIsOneOfThem()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new ServersViewModel(new LegacyServerCatalog(), new NoPicker(), "profiles.json", new UnusedSessionFactory());
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700 };
            try
            {
                window.Show();
                window.UpdateLayout();
                var view = window.GetVisualDescendants().OfType<DevicesView>().Single();
                var refresh = view.GetVisualDescendants().OfType<Button>()
                    .Single(button => AutomationProperties.GetName(button) == "Refresh devices");
                var toolbar = refresh.GetVisualAncestors().OfType<StackPanel>().First();
                var buttons = toolbar.GetVisualDescendants().OfType<Button>().Where(button => button.IsEffectivelyVisible).ToList();
                Assert.IsTrue(buttons.Count >= 2, "The devices toolbar has more than one action.");
                foreach (var button in buttons)
                {
                    Assert.IsTrue(button.Classes.Contains("icon"),
                        $"'{AutomationProperties.GetName(button)}' keeps the toolbar's icon shape.");
                    Assert.AreEqual(32, button.Bounds.Width, 0.5, "Every action of the toolbar is the same size.");
                    Assert.AreEqual(32, button.Bounds.Height, 0.5, "Every action of the toolbar is the same size.");
                }
                var hotkeys = buttons.Single(button => AutomationProperties.GetName(button) == "Device hotkeys");
                Assert.AreEqual("Hotkeys", ToolTip.GetTip(hotkeys) as string, "The button says what it opens.");
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-devices-toolbar.png"), PngBitmapEncoderOptions.Default);
                }
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, default);
    }
}
