using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Avalonia.Styling;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeviceImportViewTests
{
    [DataTestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public async Task OfflineImport_OpensSearchesSelectsChecksAcknowledgesSavesAndReconnects_AtMinimumSize(bool darkAndLongNames)
    {
        using var f = new DeviceImportTests.Fixture();
        f.Preferences.Value = new("", "", false);
        await f.LocalAsync(DeviceImportTests.Fixture.Peer, """[{"EntityId":7,"Kind":"SmartSwitch"},{"EntityId":8,"Kind":"SmartAlarm","Alias":"Oil","InGameAlarmTitle":"Oil title"},{"EntityId":9,"Kind":"SmartSwitch","Name":"Switch"}]""", 0, 0);
        if (darkAndLongNames)
        {
            var source = System.Text.Json.Nodes.JsonNode.Parse(await File.ReadAllTextAsync(f.LocalPath(DeviceImportTests.Fixture.Peer)))!;
            source["Devices"]![1]!["Alias"] = "Oil " + new string('W', 480);
            await File.WriteAllTextAsync(f.LocalPath(DeviceImportTests.Fixture.Peer), source.ToJsonString());
            f.Session.Members = [new(DeviceImportTests.Fixture.Peer, new string('W', 500))];
        }
        await UiTestSession.Current.Dispatch(async () =>
        {
            var factory = new Factory(f.Session);
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), f.Cloud.Profiles.Path, factory, deviceImporter: f.Service);
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700, RequestedThemeVariant = darkAndLongNames ? ThemeVariant.Dark : ThemeVariant.Light };
            try
            {
                await model.RefreshCommand.ExecuteAsync(null);
                window.Show();
                Assert.IsFalse(model.ImportDevicesCommand.CanExecute(null));
                model.Connection.ConnectCommand.Execute(null);
                Assert.IsTrue(model.ImportDevicesCommand.CanExecute(null));
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);
                var open = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, model.ImportDevicesCommand));
                Assert.IsTrue(open.IsEffectivelyVisible && open.IsEnabled);
                open.Command!.Execute(null);
                Task modal = model.ImportDevicesCommand.ExecutionTask!;
                var editor = model.DeviceImport!;
                Assert.IsTrue(model.HasOverlay);
                await editor.LoadCommand.ExecuteAsync(null);
                Assert.AreEqual(3, editor.Items.Count);
                Assert.IsFalse(editor.CanImport);
                Assert.IsFalse(editor.Items[0].CanChoose);
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);
                var view = window.GetVisualDescendants().OfType<DeviceImportView>().Single();
                editor.Search = "nothing matches";
                Assert.IsTrue(editor.HasNoMatches);
                editor.ClearSelectionCommand.Execute(null);
                editor.SelectAllCommand.Execute(null);
                Assert.AreEqual(2, editor.Items.Count(row => row.IsSelected));
                editor.ClearSelectionCommand.Execute(null);
                var search = view.GetVisualDescendants().OfType<TextBox>().Single(control => AutomationProperties.GetName(control) == "Search shared devices");
                search.Text = "Oil";
                editor.Items.Single(row => row.Item.Device.EntityId == 8).IsSelected = true;
                Assert.AreEqual(1, editor.VisibleItems.Count);
                Assert.IsFalse(editor.HasNoMatches);
                await editor.CheckCommand.ExecuteAsync(null);
                StringAssert.Contains(editor.VisibleItems.Single().Status, "Responding");
                view.GetVisualDescendants().OfType<CheckBox>().Single(box => ReferenceEquals(box.DataContext, editor)).IsChecked = true;
                Assert.IsTrue(editor.CanImport);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, darkAndLongNames ? "ui-device-import-dark-long.png" : "ui-device-import.png"), PngBitmapEncoderOptions.Default);
                }
                foreach (var button in view.GetVisualDescendants().OfType<Button>().Where(button => button.IsEffectivelyVisible))
                {
                    var point = button.TranslatePoint(default, window)!.Value;
                    Assert.IsTrue(point.X >= 0 && point.X + button.Bounds.Width <= window.ClientSize.Width);
                    Assert.IsTrue(point.Y >= 0 && point.Y + button.Bounds.Height <= window.ClientSize.Height);
                }
                await editor.ImportCommand.ExecuteAsync(null);
                StringAssert.Contains(editor.Message, "Imported 1");
                Assert.IsFalse(editor.CanImport);
                await editor.CloseCommand.ExecuteAsync(null);
                await modal.WaitAsync(TimeSpan.FromSeconds(3));
                Assert.IsFalse(model.HasOverlay);
                Assert.AreEqual("Oil title", model.SelectedServer!.Devices[2].InGameAlarmTitle);
                Assert.AreEqual(2, factory.Count);
                Assert.IsTrue(model.Connection.IsConnected);
                Assert.AreEqual(0, f.Cloud.Requests.Count);
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, default);
    }

    private sealed class Picker : IProfileFilePicker { public Task<string?> PickAsync() => Task.FromResult<string?>(null); }
    private sealed class Factory(DeviceImportTests.TeamSession first) : IServerSessionFactory
    {
        public int Count { get; private set; }
        public Task<IServerSession> ConnectAsync(string path, string key, CancellationToken ct)
            => Task.FromResult<IServerSession>(Count++ == 0 ? first : new DeviceImportTests.TeamSession());
    }
}
