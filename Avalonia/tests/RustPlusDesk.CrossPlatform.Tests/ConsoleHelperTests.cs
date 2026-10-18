using System.Text.Json;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Console;
using RustPlusDesk.Core.Console;
using RustPlusDesk.Infrastructure.Console;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ConsoleHelperTests
{
    private const int Bandage = -2072273936, Rock = 963906841;
    private static readonly ConsoleItemChoice[] Items = [new(Bandage, "Bandage"), new(Rock, "Rock")];

    private sealed class MemoryStore(ConsoleCommandCatalog catalog) : IConsoleCommandStore
    {
        public Dictionary<string, ConsoleCommandUserState> State { get; set; } = [];
        public int Saves { get; private set; }
        public ConsoleCommandCatalog LoadCatalog() => catalog;
        public Task<IReadOnlyDictionary<string, ConsoleCommandUserState>> LoadStateAsync(CancellationToken ct) =>
            Task.FromResult<IReadOnlyDictionary<string, ConsoleCommandUserState>>(State.ToDictionary());
        public Task SaveStateAsync(IReadOnlyDictionary<string, ConsoleCommandUserState> state, CancellationToken ct)
        {
            State = state.ToDictionary();
            Saves++;
            return Task.CompletedTask;
        }
    }

    private static ConsoleCommandCatalog Catalog() => new()
    {
        Commands =
        [
            new() { Id = "auto-run", Group = "Movement", Featured = true, Title = "Auto-Run", Description = "Run forward.", Command = "forward;sprint", Bindable = true, DefaultBind = "kp2" },
            new() { Id = "toggle-crouch", Group = "Movement", Title = "Toggle Crouch", Description = "Stay crouched.", Command = "duck", Bindable = true },
            new() { Id = "craft", Group = "Crafting", Title = "Craft Item", Description = "Queue a craft.", Command = "craft.add {item} {amount}", Bindable = true,
                Params = [new() { Name = "item", Type = "item", Default = Bandage.ToString() }, new() { Name = "amount", Type = "int", Label = "Amount", Default = "1" }] },
            new() { Id = "inertia", Title = "Head inertia", Description = "Smooth camera.", Command = "headlerp_inertia {value}",
                Params = [new() { Name = "value", Type = "bool", Default = "false" }] },
            new() { Id = "kick", Category = "admin", Group = "Players", Featured = true, Title = "Kick player", Description = "Removes a player.", Command = "global.kick \"{player}\"",
                Params = [new() { Name = "player", Type = "text" }] },
            new() { Id = "exec", Category = "admin", Title = "Meta exec", Description = "Raw quoting.", Command = "meta.exec \"a\" \"b\"", Bindable = true, RawBind = true, DefaultBind = "f5" }
        ]
    };

    [TestMethod]
    public void Rules_ResolveTemplates_AndBuildReferenceBindLines()
    {
        Assert.AreEqual("craft.add 5 2", ConsoleCommandRules.Resolve("craft.add {item} {amount}", [new("item", "5"), new("amount", "2")]));
        Assert.AreEqual("bind kp1 duck", ConsoleCommandRules.BindLine("duck", "kp1", false));
        Assert.AreEqual("bind <key> duck", ConsoleCommandRules.BindLine("duck", "", false));
        Assert.AreEqual("bind kp1 \"attack;duck\"", ConsoleCommandRules.BindLine("attack;duck", "kp1", false));
        Assert.AreEqual("bind z \"chat.say 'hi there'\"", ConsoleCommandRules.BindLine("chat.say \"hi there\"", "z", false), "Inner quotes become apostrophes.");
        Assert.AreEqual("bind z \"already\"", ConsoleCommandRules.BindLine("\"already\"", "z", false));
        Assert.AreEqual("bind f5 meta.exec \"a\" \"b\"", ConsoleCommandRules.BindLine("meta.exec \"a\" \"b\"", "f5", true));
        var spec = new ConsoleCommandSpec { Command = "forward;sprint", Group = "Movement" };
        Assert.IsTrue(ConsoleCommandRules.Matches(spec, " SPRINT ", "Auto-Run", ""));
        Assert.IsTrue(ConsoleCommandRules.Matches(spec, "movement", "Auto-Run", ""));
        Assert.IsTrue(ConsoleCommandRules.Matches(spec, "auto", "Auto-Run", ""));
        Assert.IsFalse(ConsoleCommandRules.Matches(spec, "craft", "Auto-Run", ""));
    }

    [TestMethod]
    public async Task Store_ReadsTheBundledCatalog_AndKeepsTheReferenceStateShape()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-console-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(folder, "console-helper.json");
        try
        {
            var store = new ConsoleCommandStore(path);
            var catalog = store.LoadCatalog();
            Assert.AreEqual(58, catalog.Commands.Count, "The reference catalogue ships 58 commands.");
            Assert.AreEqual(32, catalog.Commands.Count(command => command.Category == ConsoleCommandRules.Client));
            Assert.AreEqual(26, catalog.Commands.Count(command => command.Category == ConsoleCommandRules.Admin));
            Assert.AreEqual("kp1", catalog.Commands.Single(command => command.Id == "auto-attack-crouch").DefaultBind);
            Assert.AreEqual(0, (await store.LoadStateAsync(CancellationToken.None)).Count);

            await store.SaveStateAsync(new Dictionary<string, ConsoleCommandUserState>
            {
                ["craft-item"] = new() { Bind = "kp5", Values = new() { ["amount"] = "3" } }
            }, CancellationToken.None);
            using (var document = JsonDocument.Parse(await File.ReadAllTextAsync(path)))
            {
                var entry = document.RootElement.GetProperty("craft-item");
                Assert.AreEqual("kp5", entry.GetProperty("bind").GetString());
                Assert.AreEqual("3", entry.GetProperty("values").GetProperty("amount").GetString());
            }
            Assert.AreEqual("kp5", (await store.LoadStateAsync(CancellationToken.None))["craft-item"].Bind);

            await File.WriteAllTextAsync(path, "{ not json");
            var broken = new ConsoleCommandStore(path);
            Assert.AreEqual(0, (await broken.LoadStateAsync(CancellationToken.None)).Count);
            await broken.SaveStateAsync(new Dictionary<string, ConsoleCommandUserState>(), CancellationToken.None);
            Assert.AreEqual(1, Directory.GetFiles(folder, "console-helper.json.unreadable-*.bak").Length, "The unreadable file is kept as a backup.");
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public async Task ViewModel_FiltersTabsGroupsAndSearch()
    {
        var model = new ConsoleHelperViewModel(new MemoryStore(Catalog()), Items);
        await model.LoadAsync(CancellationToken.None);
        CollectionAssert.AreEqual(new[] { "Auto-Run" }, model.Featured.Select(command => command.Title).ToArray());
        CollectionAssert.AreEqual(new[] { "Crafting", "Movement", "Other" }, model.Groups.Select(group => group.Title).ToArray(), "Empty groups are listed as Other.");
        model.SearchText = "duck";
        Assert.IsFalse(model.HasFeatured);
        Assert.AreEqual("Toggle Crouch", model.Groups.Single().Commands.Single().Title);
        model.SearchText = "nothing here";
        Assert.AreEqual("No command matches your search.", model.EmptyMessage);
        model.ClearSearchCommand.Execute(null);
        model.ShowAdminCommand.Execute(null);
        Assert.AreEqual("Kick player", model.Featured.Single().Title);
        Assert.AreEqual("Meta exec", model.Groups.Single().Commands.Single().Title);

        var empty = new ConsoleHelperViewModel(new MemoryStore(new ConsoleCommandCatalog()), Items);
        await empty.LoadAsync(CancellationToken.None);
        Assert.AreEqual("Could not load the command list. See the log for details.", empty.EmptyMessage);
    }

    [TestMethod]
    public async Task ViewModel_EditsParamsAndBinds_AndSavesOnlyWhatDiffers()
    {
        var store = new MemoryStore(Catalog()) { State = new() { ["toggle-crouch"] = new() { Bind = "c" }, ["craft"] = new() { Values = new() { ["amount"] = "4" } } } };
        var model = new ConsoleHelperViewModel(store, Items);
        string? copied = null;
        model.CopyText = text => { copied = text; return Task.CompletedTask; };
        await model.LoadAsync(CancellationToken.None);
        Assert.AreEqual(0, store.Saves, "Loading never writes.");
        var crouch = model.Commands.Single(command => command.Spec.Id == "toggle-crouch");
        var craft = model.Commands.Single(command => command.Spec.Id == "craft");
        var run = model.Commands.Single(command => command.Spec.Id == "auto-run");
        Assert.AreEqual("c", crouch.BindKey);
        Assert.AreEqual($"craft.add {Bandage} 4", craft.ResolvedCommand);
        Assert.IsTrue(craft.HasChangedParams);
        Assert.AreEqual("Bandage", craft.Params[0].ItemDisplay);
        Assert.AreEqual(Bandage, craft.Params[0].ItemId);
        Assert.AreEqual("kp2", run.BindKey, "The shipped default bind applies until the user picks one.");
        Assert.AreEqual("bind kp2 \"forward;sprint\"", run.BindLine);

        craft.Params[0].ItemQuery = "r";
        Assert.IsFalse(craft.Params[0].IsItemListOpen, "Two characters are needed.");
        craft.Params[0].ItemQuery = "ro";
        Assert.AreEqual("Rock (963906841)", craft.Params[0].ItemMatches.Single().Label);
        craft.Params[0].ChooseItem(craft.Params[0].ItemMatches[0]);
        Assert.AreEqual("", craft.Params[0].ItemQuery);
        Assert.IsFalse(craft.Params[0].IsItemListOpen);
        Assert.AreEqual("bind <key> \"craft.add 963906841 4\"", craft.BindLine);
        await model.FlushAsync();
        Assert.AreEqual("963906841", store.State["craft"].Values!["item"]);
        Assert.AreEqual("4", store.State["craft"].Values!["amount"]);
        Assert.IsFalse(store.State.ContainsKey("auto-run"), "Default binds are not written.");

        await model.CopyBindCommand.ExecuteAsync(craft);
        Assert.AreEqual("Pick a key first, then the bind can be copied.", model.Notice);
        Assert.IsNull(copied);
        craft.BindKey = " kp5 ";
        await model.CopyBindCommand.ExecuteAsync(craft);
        Assert.AreEqual("bind kp5 \"craft.add 963906841 4\"", copied);
        await model.CopyCommandCommand.ExecuteAsync(crouch);
        Assert.AreEqual("duck", copied);
        await model.CopyAllBindsCommand.ExecuteAsync(null);
        Assert.AreEqual(string.Join(Environment.NewLine, "bind kp2 \"forward;sprint\"", "bind c duck", "bind kp5 \"craft.add 963906841 4\"", "bind f5 meta.exec \"a\" \"b\""), copied,
            "Every bindable command with a key, both tabs, in catalogue order.");

        craft.ResetParamsCommand.Execute(null);
        Assert.IsFalse(craft.HasChangedParams);
        var inertia = model.Commands.Single(command => command.Spec.Id == "inertia").Params[0];
        Assert.IsTrue(inertia.IsBool);
        inertia.BoolValue = true;
        Assert.AreEqual("true", inertia.Value);
        crouch.BindKey = "";
        await model.FlushAsync();
        Assert.AreEqual("kp5", store.State["craft"].Bind);
        Assert.IsNull(store.State["craft"].Values!.GetValueOrDefault("amount"), "A reset value is no longer stored.");
        Assert.AreEqual("", store.State["toggle-crouch"].Bind, "A cleared bind is remembered as cleared.");
        Assert.AreEqual("true", store.State["inertia"].Values!["value"]);

        var none = new ConsoleHelperViewModel(new MemoryStore(new() { Commands = [new() { Id = "x", Command = "x" }] }), Items) { CopyText = _ => Task.CompletedTask };
        await none.LoadAsync(CancellationToken.None);
        await none.CopyAllBindsCommand.ExecuteAsync(null);
        Assert.AreEqual("No keys bound yet. Pick a key on any command first.", none.Notice);
    }

    [TestMethod]
    public void RustKeyNames_UseRustNotation()
    {
        Assert.AreEqual("kp1", RustKeyNames.ToRustBind(Key.NumPad1, KeyModifiers.None));
        Assert.AreEqual("1", RustKeyNames.ToRustBind(Key.D1, KeyModifiers.None));
        Assert.AreEqual("=", RustKeyNames.ToRustBind(Key.OemPlus, KeyModifiers.None));
        Assert.AreEqual("q", RustKeyNames.ToRustBind(Key.Q, KeyModifiers.None));
        Assert.AreEqual("f11", RustKeyNames.ToRustBind(Key.F11, KeyModifiers.None));
        Assert.AreEqual("return", RustKeyNames.ToRustBind(Key.Enter, KeyModifiers.None));
        Assert.AreEqual("[leftcontrol+leftshift+f11]", RustKeyNames.ToRustBind(Key.F11, KeyModifiers.Shift | KeyModifiers.Control));
        Assert.AreEqual("leftshift", RustKeyNames.ToRustBind(Key.LeftShift, KeyModifiers.Shift), "A modifier alone has no prefix.");
        Assert.IsNull(RustKeyNames.ToRustBind(Key.LWin, KeyModifiers.None));
        Assert.IsNull(RustKeyNames.ToRustBind(Key.VolumeUp, KeyModifiers.None), "No guess for unknown keys.");
        CollectionAssert.Contains(RustKeyNames.MouseAndWheel.ToArray(), "mousewheeldown");
    }

    [TestMethod]
    public async Task KeyCapture_RecordsRustNames_AndRejectsUnknownKeys()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var window = new KeyCaptureWindow("kp1");
            try
            {
                window.Show();
                Assert.AreEqual("kp1", window.CapturedKey);
                window.OnKeyDown(window, new KeyEventArgs { Key = Key.VolumeUp, RoutedEvent = InputElement.KeyDownEvent });
                Assert.AreEqual("kp1", window.CapturedKey);
                Assert.IsTrue(window.GetVisualDescendants().OfType<TextBlock>().Any(text => text.Text == "That key has no Rust equivalent. Pick another one." && text.IsVisible));
                window.OnKeyDown(window, new KeyEventArgs { Key = Key.G, KeyModifiers = KeyModifiers.Alt, RoutedEvent = InputElement.KeyDownEvent });
                Assert.AreEqual("[leftalt+g]", window.CapturedKey);
                var mouse = window.GetVisualDescendants().OfType<ComboBox>().Single();
                mouse.SelectedItem = "mouse4";
                Assert.AreEqual("mouse4", window.CapturedKey);
                var clear = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Clear key");
                clear.RaiseEvent(new global::Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                Assert.IsNull(window.CapturedKey);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }

    [TestMethod]
    public async Task ConsoleHelperView_RendersCommandCards()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new ConsoleHelperViewModel(new ConsoleCommandStore(Path.Combine(Path.GetTempPath(), "rpd-missing-" + Guid.NewGuid().ToString("N"), "c.json")),
                RustPlusDesk.Infrastructure.Items.BundledItemCatalog.Icons.Values.Select(item => new ConsoleItemChoice(item.Id, item.Name)));
            await model.LoadAsync(CancellationToken.None);
            var window = new Window { Width = 1000, Height = 700, Content = new Border { Padding = new global::Avalonia.Thickness(16), Child = new ConsoleHelperView { DataContext = model } } };
            try
            {
                window.Show();
                window.UpdateLayout();
                Assert.IsTrue(model.Featured.Count > 0);
                var copy = window.GetVisualDescendants().OfType<Button>().First(button => AutomationProperties.GetName(button) == "Copy command");
                Assert.IsNotNull(copy.Command);
                Assert.IsTrue(window.GetVisualDescendants().OfType<Button>().Any(button => AutomationProperties.GetName(button) == "Pop out console helper" && button.IsEffectivelyVisible));
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-console-helper.png"), PngBitmapEncoderOptions.Default);
                }
                model.ShowAdminCommand.Execute(null);
                window.UpdateLayout();
                Assert.IsTrue(window.GetVisualDescendants().OfType<TextBlock>().Any(text => text.Text?.StartsWith("These only work if you are an owner") == true && text.IsEffectivelyVisible));
            }
            finally { window.Close(); }
            return true;
        }, default);
    }

    [TestMethod]
    public async Task Shell_OpensTheConsoleTool_AndPopOutReturnsToTheMap()
    {
        var model = new RustPlusDesk.Avalonia.Features.Servers.ServersViewModel(new LegacyServerCatalogStub(), new PickerStub(), "profiles.json", new UnusedSessionFactory(),
            consoleStore: new MemoryStore(Catalog()));
        Assert.IsTrue(model.HasConsoleTool);
        int popOuts = 0;
        model.ConsolePopOutRequested += (_, _) => popOuts++;
        model.OpenConsoleCommand.Execute(null);
        Assert.IsTrue(model.IsConsoleSelected);
        Assert.IsNotNull(model.ConsoleHelper);
        await CompanionListenerViewModelTests.Eventually(() => model.ConsoleHelper!.Commands.Count == 6);
        model.ConsoleHelper!.PopOutCommand.Execute(null);
        Assert.AreEqual(1, popOuts);
        Assert.IsFalse(model.IsToolSelected, "Popping out returns the main window to the map.");
        Assert.AreSame(model.ConsoleHelper, model.EnsureConsoleHelper(), "The popout shares the same model.");
        Assert.IsFalse(new RustPlusDesk.Avalonia.Features.Servers.ServersViewModel(new LegacyServerCatalogStub(), new PickerStub(), "profiles.json", new UnusedSessionFactory()).HasConsoleTool);
    }

    private sealed class LegacyServerCatalogStub : RustPlusDesk.Core.Servers.IServerCatalog
    {
        public Task<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>>([]);
    }

    private sealed class PickerStub : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
