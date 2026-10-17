using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Core.Devices;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeviceImportViewModelTests
{
    [TestMethod]
    public async Task Closing_CancelsAndAwaitsPreviewCleanup_AndNeverSignalsASave()
    {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var model = new DeviceImportViewModel(async ct =>
        {
            started.TrySetResult();
            try { await Task.Delay(Timeout.Infinite, ct); }
            finally { await release.Task; }
            throw new AssertFailedException("Canceled preview must not complete.");
        }, (_, _, _, _) => throw new AssertFailedException("No check requested."), (_, _, _, _, _) => throw new AssertFailedException("No save requested."));
        Task loading = model.LoadCommand.ExecuteAsync(null);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(3));
        Task close = model.CloseCommand.ExecuteAsync(null);
        Assert.IsFalse(close.IsCompleted);
        Assert.IsFalse(model.Completion.IsCompleted);
        release.TrySetResult();
        await Task.WhenAll(loading, close).WaitAsync(TimeSpan.FromSeconds(3));
        Assert.IsNull(await model.Completion);
        Assert.IsFalse(model.CanImport);
    }
}
