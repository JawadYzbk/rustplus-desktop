using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Toasts;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ToastTests
{
    private sealed class ManualScheduler
    {
        public List<(TimeSpan Delay, Action Run, Handle Handle)> Pending { get; } = [];
        public IDisposable Schedule(TimeSpan delay, Action run)
        {
            var handle = new Handle();
            Pending.Add((delay, run, handle));
            return handle;
        }
        public void Fire(int index) { var entry = Pending[index]; if (!entry.Handle.Disposed) entry.Run(); }
    }

    private sealed class Handle : IDisposable
    {
        public bool Disposed { get; private set; }
        public void Dispose() => Disposed = true;
    }

    [TestMethod]
    public void Toasts_TimeOutStackToFiveAndCloseOnce()
    {
        var scheduler = new ManualScheduler();
        var model = new ToastsViewModel(scheduler.Schedule);
        int closed = 0;
        var first = new ToastItem { Title = "Raid", Message = "Base alarm", Kind = ToastKind.Danger, Closed = () => closed++ };
        model.Add(first);
        Assert.AreEqual(TimeSpan.FromSeconds(8), scheduler.Pending.Single().Delay, "Reference default: eight seconds.");
        scheduler.Fire(0);
        Assert.AreEqual(0, model.Items.Count);
        Assert.AreEqual(1, closed);
        model.Dismiss(first);
        Assert.AreEqual(1, closed, "Closed runs exactly once.");

        var pinned = new ToastItem { Message = "Pinned", Timeout = TimeSpan.Zero };
        model.Add(pinned);
        Assert.AreEqual(1, scheduler.Pending.Count, "Pinned toasts never time out.");
        Assert.IsFalse(pinned.HasTitle);
        for (int i = 0; i < 5; i++) model.Show($"T{i}", "m", ToastKind.Success);
        Assert.AreEqual(ToastsViewModel.MaxVisible, model.Items.Count);
        Assert.IsFalse(model.Items.Contains(pinned), "The oldest is evicted past five.");
        Assert.AreEqual("T4", model.Items[^1].Title, "Newest last, nearest the corner.");
        Assert.IsTrue(model.Items[0].IsSuccess);
        var evicted = model.Items[0];
        model.DismissCommand.Execute(evicted);
        Assert.IsNull(evicted.Timer, "A dismissed toast's timer is cancelled.");
        Assert.AreEqual(2, scheduler.Pending.Count(entry => entry.Handle.Disposed), "The fired and the dismissed timers are released.");
        Assert.AreEqual(4, model.Items.Count);
        Assert.AreEqual(RustPlusDesk.Avalonia.Shared.Icons.FluentSymbol.Warning24, ToastsViewModel.SymbolFor(ToastKind.Caution));
    }

    [TestMethod]
    public async Task Notifications_RaiseToastsUnlessDisabled() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var shell = new ServersViewModel(new CatalogStub(), new PickerStub(), "profiles.json", new UnusedSessionFactory(), appSettings: settings);
        shell.Notifications.Handle(new(CompanionUpdateKind.Alarm, "Raid", "Base alarm") { Host = "1.2.3.4", Port = 1 });
        var toast = shell.Toasts.Items.Single();
        Assert.AreEqual("Raid", toast.Title);
        Assert.IsTrue(toast.IsDanger);
        shell.Notifications.Handle(new(CompanionUpdateKind.Death, "You were killed by Ann", "Ann") { ServerName = "Rusty" });
        Assert.IsTrue(shell.Toasts.Items[^1].IsCaution);
        settings.Set(SettingKeys.NotificationsToastEnabled, false);
        shell.Notifications.Handle(new(CompanionUpdateKind.Pairing, "Switch", "x") { EntityId = 3 });
        Assert.AreEqual(2, shell.Toasts.Items.Count, "NotificationsToastEnabled off: rows still arrive, no toast.");
        Assert.AreEqual(3, shell.Notifications.Rows.Count);
        await Task.CompletedTask;
        return true;
    }, default);

    [TestMethod]
    public async Task ToastHost_RendersStackedToasts()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new ToastsViewModel((_, _) => new Handle());
            model.Show("HV Rockets Raid wake up", "Your base is under attack!", ToastKind.Danger);
            model.Show("Server muted", "Notifications from Test Server are now muted.");
            model.Show("Smart Switch", "Device paired (ID 42)", ToastKind.Success);
            var window = new Window { Width = 1000, Height = 700, Content = new ToastHost { DataContext = model, Margin = new global::Avalonia.Thickness(16) } };
            try
            {
                window.Show();
                window.UpdateLayout();
                Assert.AreEqual(3, window.GetVisualDescendants().OfType<Border>().Count(border => border.Classes.Contains("toast")));
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-toasts.png"), PngBitmapEncoderOptions.Default);
                }
                var dismiss = window.GetVisualDescendants().OfType<Button>().First(button => global::Avalonia.Automation.AutomationProperties.GetName(button) == "Dismiss toast");
                dismiss.Command!.Execute(dismiss.CommandParameter);
                Assert.AreEqual(2, model.Items.Count);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }

    private sealed class CatalogStub : RustPlusDesk.Core.Servers.IServerCatalog
    {
        public Task<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>>([]);
    }

    private sealed class PickerStub : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
