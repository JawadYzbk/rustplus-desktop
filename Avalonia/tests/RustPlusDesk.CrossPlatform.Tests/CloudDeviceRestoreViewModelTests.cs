using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudDeviceRestoreViewModelTests
{
    [TestMethod]
    public async Task PreviewIsReadOnly_RequiresAcknowledgment_AndRestoreKeepsLatestRevision()
    {
        var service = new Service();
        var model = new CloudDeviceRestoreViewModel("fixture", service.Server, "user", "steam", service);
        Assert.IsFalse(model.RestoreCommand.CanExecute(null));
        await model.PreviewCommand.ExecuteAsync(null);
        Assert.AreEqual(1, model.Devices.Count);
        Assert.IsFalse(model.RestoreCommand.CanExecute(null));
        Assert.AreEqual(0, service.Restores);
        model.Acknowledged = true;
        await model.RestoreCommand.ExecuteAsync(null);
        Assert.AreEqual(1, service.Restores);
        Assert.IsNull(model.Snapshot);
        Assert.IsFalse(model.Acknowledged);
        await model.PreviewCommand.ExecuteAsync(null);
        Assert.AreEqual("saved", service.LastServer!.SourceRevision);
        await model.CloseCommand.ExecuteAsync(null);
        Assert.IsTrue(await model.Completion);
    }

    [TestMethod]
    public async Task PreviousWipeOrEmptySnapshot_CannotRestore_EvenWhenAcknowledged()
    {
        var service = new Service();
        var model = new CloudDeviceRestoreViewModel("fixture", service.Server, "user", "steam", service);
        service.Snapshot = service.Snapshot with { PredatesWipe = true };
        await model.PreviewCommand.ExecuteAsync(null);
        model.Acknowledged = true;
        Assert.IsFalse(model.RestoreCommand.CanExecute(null));
        StringAssert.Contains(model.Message, "predates");
        service.Snapshot = service.Snapshot with { PredatesWipe = false, Devices = [] };
        await model.PreviewCommand.ExecuteAsync(null);
        model.Acknowledged = true;
        Assert.IsFalse(model.RestoreCommand.CanExecute(null));
        await model.CloseCommand.ExecuteAsync(null);
        Assert.IsFalse(await model.Completion);
    }

    [DataTestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public async Task CloseCancelsAndAwaitsOwnedOperation_BeforeCompleting(bool restoring)
    {
        var service = new Service();
        var model = new CloudDeviceRestoreViewModel("fixture", service.Server, "user", "steam", service);
        if (restoring) { await model.PreviewCommand.ExecuteAsync(null); model.Acknowledged = true; }
        service.Wait = true;
        var operation = restoring ? model.RestoreCommand.ExecuteAsync(null) : model.PreviewCommand.ExecuteAsync(null);
        await service.Started.Task.WaitAsync(TimeSpan.FromSeconds(3));
        var close = model.CloseCommand.ExecuteAsync(null);
        Assert.IsFalse(close.IsCompleted);
        Assert.IsFalse(model.Completion.IsCompleted);
        service.Release.TrySetResult();
        await Task.WhenAll(operation, close).WaitAsync(TimeSpan.FromSeconds(3));
        Assert.IsFalse(await model.Completion);
        Assert.IsFalse(model.IsBusy);
        Assert.IsNull(model.Snapshot);
    }

    [TestMethod]
    public async Task FailuresClearOldSnapshot_AndNeverShowExceptionDetails()
    {
        var service = new Service();
        var model = new CloudDeviceRestoreViewModel("fixture", service.Server, "user", "steam", service);
        await model.PreviewCommand.ExecuteAsync(null);
        model.Acknowledged = true;
        service.Fail = true;
        await model.RestoreCommand.ExecuteAsync(null);
        Assert.IsNull(model.Snapshot);
        Assert.IsFalse(model.Acknowledged);
        Assert.IsFalse(model.Message.Contains("synthetic-secret"));
        Assert.IsFalse(model.RestoreCommand.CanExecute(null));
    }

    internal sealed class Service : ICloudDeviceRestore
    {
        public ServerSummary Server { get; } = new("Fixture server", "fixture.test", 28082) { ProfileKey = "profile", SourceRevision = "original" };
        public CloudDeviceSnapshot Snapshot { get; set; } = new("user", "steam", "profile", "snapshot", null, false,
            [new(8, "SmartAlarm", "Fixture alarm", null, false, [], InGameAlarmTitle: "Fixture title")]);
        public ServerSummary? LastServer { get; private set; }
        public int Restores { get; private set; }
        public bool Wait { get; set; }
        public bool Fail { get; set; }
        public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private async Task WaitAsync(CancellationToken ct)
        {
            if (Fail) throw new IOException("synthetic-secret-credentials-and-path");
            if (!Wait) return;
            Started.TrySetResult();
            await Release.Task;
            ct.ThrowIfCancellationRequested();
        }
        public async Task<CloudDeviceSnapshot> PreviewAsync(string path, ServerSummary server, string userId, string steamId, CancellationToken ct)
        { LastServer = server; await WaitAsync(ct); return Snapshot; }
        public async Task<CloudDeviceRestoreResult> RestoreAsync(string path, ServerSummary server, CloudDeviceSnapshot preview, CancellationToken ct)
        { await WaitAsync(ct); Restores++; return new(server with { SourceRevision = "saved" }, 1); }
    }
}
