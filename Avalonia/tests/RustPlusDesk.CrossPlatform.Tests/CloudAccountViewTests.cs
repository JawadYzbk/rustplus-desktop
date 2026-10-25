using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Accounts;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudAccountViewTests
{
    [DataTestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public async Task AccountPanel_RendersCloudAndCompanionAtMinimumSizeAndKeepsCloseVisible(bool light)
    {
        await using var session = HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder));
        await session.Dispatch(async () =>
        {
            var cloud = new Service();
            var model = new ServersViewModel(new Catalog(), new Picker(), "fixture-profiles.json", new UnusedSessionFactory(),
                accountService: new CompanionAccountViewModelTests.Service(), cloudAccountService: cloud);
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700, RequestedThemeVariant = light ? ThemeVariant.Light : ThemeVariant.Dark };
            try
            {
                window.Show(); model.OpenAccountCommand.Execute(null);
                Assert.IsNotNull(model.CloudAccount);
                await model.CloudAccount.CheckCommand.ExecuteAsync(null);
                Assert.AreEqual(0, cloud.SignIns);
                await model.CloudAccount.SignInCommand.ExecuteAsync(null);
                Assert.AreEqual(1, cloud.SignIns);
                using var frame = window.CaptureRenderedFrame();
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, $"ui-cloud-account-{(light ? "light" : "dark")}.png"), PngBitmapEncoderOptions.Default);
                var view = window.GetVisualDescendants().OfType<CloudAccountView>().Single();
                Assert.IsTrue(view.IsEffectivelyVisible);
                Assert.IsTrue(view.GetVisualDescendants().OfType<TextBlock>().Any(t => t.Text == "Fixture cloud account"));
                var close = window.GetVisualDescendants().OfType<Button>().Single(b => ReferenceEquals(b.Command, model.CloseAccountCommand));
                var origin = close.TranslatePoint(default, window)!.Value;
                Assert.IsTrue(origin.Y >= 0 && origin.Y + close.Bounds.Height <= window.ClientSize.Height);
                var scroll = view.GetVisualAncestors().OfType<ScrollViewer>().First();
                scroll.ScrollToEnd();
                using (var bottom = window.CaptureRenderedFrame()) Assert.IsNotNull(bottom);
                await model.CloudAccount.SignOutCommand.ExecuteAsync(null);
                Assert.AreEqual(CloudAccountState.SignedOut, model.CloudAccount.Account.State);
                await model.CloseAccountCommand.ExecuteAsync(null);
                Assert.IsFalse(model.HasOverlay);
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [TestMethod]
    public async Task CancelAndShutdown_AwaitOwnedOperation_AndNeverExposeRawErrors()
    {
        var service = new Service();
        var model = new CloudAccountViewModel(service);
        service.WaitForCancellation = true;
        var running = model.SignInCommand.ExecuteAsync(null);
        await service.Started.Task;
        Assert.IsFalse(model.CheckCommand.CanExecute(null));
        await model.StopAsync(shutdown: true);
        await running;
        Assert.IsTrue(service.Canceled);
        Assert.IsFalse(model.IsBusy);
        Assert.IsFalse(model.SignInCommand.CanExecute(null));
        Assert.IsFalse(model.Message.Contains("synthetic-secret"));
    }

    [TestMethod]
    public async Task AccountClose_CancelsCloudLoginWithoutStoppingCompanionIdentity()
    {
        var service = new Service { WaitForCancellation = true };
        var model = new ServersViewModel(new Catalog(), new Picker(), "fixture-profiles.json", new UnusedSessionFactory(), cloudAccountService: service);
        model.OpenAccountCommand.Execute(null);
        var login = model.CloudAccount!.SignInCommand.ExecuteAsync(null);
        await service.Started.Task;
        await model.CloseAccountCommand.ExecuteAsync(null);
        await login;
        Assert.IsTrue(service.Canceled);
        Assert.IsFalse(model.IsAccountOpen);
        await model.ShutdownAsync();
    }

    private sealed class Service : ICloudAccountService
    {
        public int SignIns;
        public bool WaitForCancellation, Canceled;
        public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task<CloudAccount> CheckAsync(CancellationToken ct) => Task.FromResult(new CloudAccount(CloudAccountState.SignedOut));
        public async Task<CloudAccount> SignInAsync(CancellationToken ct)
        {
            SignIns++; Started.TrySetResult();
            if (WaitForCancellation)
                try { await Task.Delay(Timeout.Infinite, ct); }
                catch (OperationCanceledException) { Canceled = true; throw; }
            return new(CloudAccountState.Active, "fixture-cloud-id", "Fixture cloud account", "fixture@example.test");
        }
        public Task SignOutAsync(CancellationToken ct) => Task.CompletedTask;
    }
    private sealed class Catalog : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken ct) => Task.FromResult<IReadOnlyList<ServerSummary>>([]);
    }
    private sealed class Picker : IProfileFilePicker { public Task<string?> PickAsync() => Task.FromResult<string?>(null); }
}
