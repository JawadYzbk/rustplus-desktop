using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Logic;

namespace RustPlusDesk.Avalonia.Features.Logic;

/// <summary>
/// Reference the logic engine's view of the app: the paired devices, the switch it drives, the countdown it starts
/// and the refresh an availability check needs. The engine itself stays free of the app's types.
/// </summary>
public sealed class LogicEngineHost(DevicesViewModel devices, Func<Task> refreshDevices,
    Func<LogicStep, CancellationToken, Task<bool>> startTimer) : ILogicEngineHost
{
    private readonly DevicesViewModel devices = devices;
    private readonly Func<Task> refreshDevices = refreshDevices;
    private readonly Func<LogicStep, CancellationToken, Task<bool>> startTimer = startTimer;

    public LogicDeviceState? Device(uint entityId)
    {
        var row = Rows().FirstOrDefault(item => item.Definition.EntityId == entityId);
        if (row is null) return null;
        return new LogicDeviceState(row.Definition.EntityId, row.IsOn, row.IsMissing, row.Name);
    }

    public IReadOnlyList<LogicDeviceState> GroupSwitches(string groupName) => Rows()
        .Where(row => row.Definition.Kind == DeviceKind.Group && row.Name == groupName)
        .SelectMany(row => row.DescendantsAndSelf())
        .Where(row => row.Definition.Kind == DeviceKind.Switch)
        .Select(row => new LogicDeviceState(row.Definition.EntityId, row.IsOn, row.IsMissing, row.Name))
        .ToArray();

    public async Task<bool> SetSwitchAsync(uint entityId, bool value, CancellationToken ct) =>
        await devices.SetSwitchStateAsync(entityId, value, ct).ConfigureAwait(false);

    public Task RefreshDevicesAsync(CancellationToken ct)
    {
        refreshDevices().GetAwaiter().GetResult();
        return Task.CompletedTask;
    }

    public Task<bool> StartTimerAsync(LogicStep step, CancellationToken ct) => startTimer(step, ct);

    private IEnumerable<DeviceRowViewModel> Rows() =>
        devices.Devices.SelectMany(row => row.DescendantsAndSelf());
}
