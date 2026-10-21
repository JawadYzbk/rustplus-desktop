using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeviceOrganizerViewModelTests
{
    [TestMethod]
    public async Task IconSearchAndSelection_AreExplicit_AndSaveResetKeepTheSelectedDevice()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var choices = RustPlusDesk.Infrastructure.Items.BundledItemCatalog.Icons.Values.ToArray();
        var model = new DeviceOrganizerViewModel(file.Path, await file.ServerAsync(), new LegacyDeviceProfileEditor(), choices);
        model.Selected = model.Devices[0].Children[0];
        model.IconSearch = "smart.switch";
        Assert.AreEqual(1, model.IconChoices.Count());
        model.SelectedIcon = model.IconChoices.Single();
        Assert.IsNull((await file.ServerAsync()).Devices[0].Children[0].CustomIconId);
        await model.SaveIconCommand.ExecuteAsync(null);
        Assert.AreEqual("0/0", model.Selected!.DeviceKey);
        Assert.AreEqual(988652725, model.Selected.CustomIconId);
        Assert.AreEqual(model.Selected.CustomIconId, model.SelectedIcon!.Id);
        model.IconSearch = "-695978112";
        Assert.AreEqual("smart.alarm", model.IconChoices.Single().ShortName);
        await model.ResetIconCommand.ExecuteAsync(null);
        Assert.IsNull(model.Selected!.CustomIconId);
        Assert.IsNull(model.SelectedIcon);
        Assert.IsFalse(model.SaveIconCommand.CanExecute(null));
        await model.CloseCommand.ExecuteAsync(null);
        Assert.IsNotNull(await model.Completion);
    }

    [TestMethod]
    public async Task Rename_Group_MoveAndReorder_KeepSelectionAndUseFreshRevisions()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var model = new DeviceOrganizerViewModel(file.Path, await file.ServerAsync(), new LegacyDeviceProfileEditor());
        model.Selected = model.Devices[0].Children[0];
        Assert.AreEqual("Alias", model.Alias);
        model.Alias = "Renamed";
        await model.RenameCommand.ExecuteAsync(null);
        Assert.AreEqual("Renamed", model.Selected!.Name);
        model.GroupName = "Inner";
        await model.GroupCommand.ExecuteAsync(null);
        Assert.AreEqual("Inner", model.Selected!.Name);
        Assert.IsFalse(model.Destinations.Any(item => item.DeviceKey == model.Selected.DeviceKey));
        Assert.IsFalse(model.Destinations.Any(item => item.DeviceKey?.StartsWith(model.Selected.DeviceKey + "/", StringComparison.Ordinal) == true));
        model.Destination = model.Destinations[0];
        await model.MoveCommand.ExecuteAsync(null);
        // Moving the only child out prunes the emptied "Saved group", as the reference drop does.
        Assert.AreEqual("1", model.Selected!.DeviceKey);
        Assert.AreEqual(2, model.Devices.Count);
        Assert.IsFalse(model.MoveDownCommand.CanExecute(null));
        await model.MoveUpCommand.ExecuteAsync(null);
        Assert.AreEqual("0", model.Selected!.DeviceKey);
        await model.MoveDownCommand.ExecuteAsync(null);
        Assert.AreEqual("1", model.Selected!.DeviceKey);
        Assert.IsFalse(model.Completion.IsCompleted);
        await model.CloseCommand.ExecuteAsync(null);
        var saved = await model.Completion;
        Assert.IsNotNull(saved);
        Assert.AreEqual("Renamed", saved.Devices[1].Children[0].Name);
    }

    [TestMethod]
    public async Task Delete_RequiresConfirmation_SelectsNeighbour_AndSelectionChangeCancelsConfirmation()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var model = new DeviceOrganizerViewModel(file.Path, await file.ServerAsync(), new LegacyDeviceProfileEditor());
        model.Selected = model.Devices[0];
        StringAssert.Contains(model.DeleteConfirmation, "1 device(s)");
        await model.DeleteCommand.ExecuteAsync(null);
        Assert.IsTrue(model.IsDeleteConfirming);
        Assert.AreEqual(1, Directory.GetFiles(file.Directory).Length);
        model.Selected = model.Devices[1];
        Assert.IsFalse(model.IsDeleteConfirming);
        model.Selected = model.Devices[0];
        await model.DeleteCommand.ExecuteAsync(null);
        model.CancelDeleteCommand.Execute(null);
        Assert.IsFalse(model.IsDeleteConfirming);
        Assert.AreEqual(2, model.Devices.Count);
        await model.DeleteCommand.ExecuteAsync(null);
        await model.DeleteCommand.ExecuteAsync(null);
        Assert.AreEqual(1, model.Devices.Count);
        Assert.AreEqual("Duplicate", model.Selected!.Name);
        Assert.AreEqual("0", model.Selected.DeviceKey);
        StringAssert.Contains(model.Message, "Deleted 1 device(s)");
        await model.DeleteCommand.ExecuteAsync(null);
        await model.DeleteCommand.ExecuteAsync(null);
        Assert.AreEqual(0, model.Devices.Count);
        Assert.IsNull(model.Selected);
        Assert.IsFalse(model.DeleteCommand.CanExecute(null));
        await model.CloseCommand.ExecuteAsync(null);
        Assert.AreEqual(0, (await model.Completion)!.Devices.Count);
    }

    [TestMethod]
    public async Task CloseWithoutChange_DoesNotRequestReloadOrCreateBackup()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var model = new DeviceOrganizerViewModel(file.Path, await file.ServerAsync(), new LegacyDeviceProfileEditor());
        await model.RenameCommand.ExecuteAsync(null);
        await model.CloseCommand.ExecuteAsync(null);
        Assert.IsNull(await model.Completion);
        Assert.AreEqual(1, Directory.GetFiles(file.Directory).Length);
    }

    [TestMethod]
    public async Task ChangedRevision_DisablesFurtherWrites_AndDoesNotExposeExceptionData()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var original = await file.ServerAsync();
        var model = new DeviceOrganizerViewModel(file.Path, original, new LegacyDeviceProfileEditor());
        string changed = await File.ReadAllTextAsync(file.Path) + "\n";
        await File.WriteAllTextAsync(file.Path, changed);
        model.Alias = "Changed";
        await model.RenameCommand.ExecuteAsync(null);
        StringAssert.Contains(model.Message, "file changed");
        Assert.IsFalse(model.RenameCommand.CanExecute(null));
        Assert.IsFalse(model.GroupCommand.CanExecute(null));
        Assert.IsFalse(model.Message.Contains(file.Path, StringComparison.Ordinal));
        Assert.AreEqual(changed, await File.ReadAllTextAsync(file.Path));
        await model.CloseCommand.ExecuteAsync(null);
        Assert.IsNull(await model.Completion);
    }

    [TestMethod]
    public async Task CloseDuringSave_CancelsAndAwaitsWriter_BeforeCompleting()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var editor = new WaitingEditor();
        var model = new DeviceOrganizerViewModel(file.Path, await file.ServerAsync(), editor);
        Task saving = model.RenameCommand.ExecuteAsync(null);
        await editor.Started.Task;
        Assert.IsFalse(model.GroupCommand.CanExecute(null));
        Task closing = model.CloseCommand.ExecuteAsync(null);
        Assert.IsFalse(closing.IsCompleted);
        Assert.IsFalse(model.Completion.IsCompleted);
        editor.Release.TrySetResult();
        await Task.WhenAll(saving, closing).WaitAsync(TimeSpan.FromSeconds(4));
        Assert.IsTrue(editor.Canceled);
        Assert.IsNull(await model.Completion);
        Assert.AreEqual(1, Directory.GetFiles(file.Directory).Length);
    }

    [TestMethod]
    public async Task RootModal_BlocksOtherLoadsAndShutdown_AwaitsOwnedWrite()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var editor = new WaitingEditor();
        var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), file.Path, new UnusedSessionFactory(), deviceEditor: editor);
        await model.RefreshCommand.ExecuteAsync(null);
        Task organizing = model.OrganizeDevicesCommand.ExecuteAsync(null);
        Assert.IsTrue(model.HasOverlay);
        Assert.IsFalse(model.RefreshCommand.CanExecute(null));
        Assert.IsFalse(model.OpenFileCommand.CanExecute(null));
        Task saving = model.Organizer!.RenameCommand.ExecuteAsync(null);
        await editor.Started.Task;
        Task shutdown = model.ShutdownAsync();
        Assert.IsFalse(shutdown.IsCompleted);
        editor.Release.TrySetResult();
        await Task.WhenAll(saving, organizing, shutdown).WaitAsync(TimeSpan.FromSeconds(4));
        Assert.IsTrue(editor.Canceled);
        Assert.IsFalse(model.HasOrganizer);
    }

    [TestMethod]
    public async Task SavedOrganization_ReloadsSelectedTree_AndRestartsOnlyThePreviouslyConnectedSession()
    {
        using var file = new DeviceProfileEditorTests.Fixture();
        var sessions = new Sessions();
        var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), file.Path, sessions, deviceEditor: new LegacyDeviceProfileEditor());
        try
        {
            await model.RefreshCommand.ExecuteAsync(null);
            model.Connection.ConnectCommand.Execute(null);
            await CompanionListenerViewModelTests.Eventually(() => model.Connection.IsConnected);
            string key = model.SelectedServer!.ProfileKey;
            Task organizing = model.OrganizeDevicesCommand.ExecuteAsync(null);
            model.Organizer!.Selected = model.Organizer.Devices[0].Children[0];
            model.Organizer.Alias = "New label";
            await model.Organizer.RenameCommand.ExecuteAsync(null);
            await model.Organizer.CloseCommand.ExecuteAsync(null);
            await organizing.WaitAsync(TimeSpan.FromSeconds(4));
            await CompanionListenerViewModelTests.Eventually(() => sessions.Opened.Count == 2 && model.Connection.IsConnected);
            Assert.IsTrue(sessions.Opened[0].Disposed);
            Assert.AreEqual(key, model.SelectedServer!.ProfileKey);
            Assert.AreEqual("New label", model.SelectedServer.Devices[0].Children[0].Name);
        }
        finally { await model.ShutdownAsync(); }
    }

    private sealed class Sessions : IServerSessionFactory
    {
        public List<Session> Opened { get; } = [];
        public Task<IServerSession> ConnectAsync(string profilesPath, string profileKey, CancellationToken cancellationToken)
        { var session = new Session(); Opened.Add(session); return Task.FromResult<IServerSession>(session); }
    }
    private sealed class Session : IServerSession
    {
        public IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Fixture", 1, 100, 0);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public bool Disposed { get; private set; }
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public ValueTask DisposeAsync() { Disposed = true; return ValueTask.CompletedTask; }
    }

    private sealed class Picker : IProfileFilePicker
    {
        public Task<string?> PickAsync() => throw new AssertFailedException("Picker must remain blocked.");
    }

    private sealed class WaitingEditor : IDeviceProfileEditor
    {
        public Task<DeviceEditResult> SaveIconAsync(string path, ServerSummary original, string deviceKey, int? iconId, string? shortName, CancellationToken ct) => throw new AssertFailedException();
        public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public bool Canceled { get; private set; }
        public async Task<DeviceEditResult> RenameAsync(string path, ServerSummary original, string deviceKey, string alias, CancellationToken ct)
        {
            Started.TrySetResult();
            await Release.Task;
            Canceled = ct.IsCancellationRequested;
            ct.ThrowIfCancellationRequested();
            throw new AssertFailedException("Test must cancel the pending writer.");
        }
        public Task<DeviceEditResult> GroupAsync(string path, ServerSummary original, string deviceKey, string name, CancellationToken ct) => throw new AssertFailedException();
        public Task<DeviceEditResult> SaveAlarmTitleAsync(string path, ServerSummary original, string deviceKey, string title, CancellationToken ct) => throw new AssertFailedException();
        public Task<DeviceEditResult> SaveAlarmAudioAsync(string path, ServerSummary original, string deviceKey, string? audioFilePath, CancellationToken ct) => throw new AssertFailedException();
        public Task<DeviceEditResult> MoveAsync(string path, ServerSummary original, string deviceKey, string? parentKey, int? index, CancellationToken ct) => throw new AssertFailedException();
        public Task<DeviceDeleteResult> DeleteAsync(string path, ServerSummary original, string deviceKey, CancellationToken ct) => throw new AssertFailedException();
    }
}
