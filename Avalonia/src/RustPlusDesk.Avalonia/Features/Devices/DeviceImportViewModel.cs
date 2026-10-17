using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Devices;

namespace RustPlusDesk.Avalonia.Features.Devices;

public sealed partial class DeviceImportRowViewModel(DeviceImportCandidate item) : ObservableObject
{
    public DeviceImportCandidate Item { get; } = item;
    public string Name => Item.Device.DisplayName;
    public string DefaultIconShortName => Item.Device.Kind?.Replace(" ", "", StringComparison.Ordinal).ToLowerInvariant() switch
    {
        "smartswitch" => "smart.switch",
        "smartalarm" => "smart.alarm",
        "storagemonitor" => "storage.monitor",
        _ => "paper"
    };
    public string Owner => $"{Item.Device.Kind} · {Item.OwnerName} · {Item.OwnerSteamId} · {(Item.Source == DeviceImportSource.Cloud ? "Cloud" : "Local snapshot")}";
    public bool CanChoose => Item.State == DeviceImportState.Unchecked;
    [ObservableProperty] private bool isSelected = item.State == DeviceImportState.Unchecked;
    [ObservableProperty]
    private string status = item.State switch
    {
        DeviceImportState.AlreadyPresent => "Already in your saved device tree",
        DeviceImportState.PreviousWipe => "Previous wipe — cannot import",
        DeviceImportState.Unsupported => "Unsupported device kind",
        _ => item.UnknownWipe ? "Unknown wipe history — review IDs in Rust" : "Not checked on the live server yet"
    };
}

/// <summary>Explicit search/selection/check/save workflow; preview data is never trusted as write authority.</summary>
public sealed partial class DeviceImportViewModel(
    Func<CancellationToken, Task<DeviceImportPreview>> load,
    Func<DeviceImportPreview, IReadOnlySet<uint>, IProgress<DeviceImportProgress>, CancellationToken, Task<DeviceImportCheck>> check,
    Func<DeviceImportPreview, IReadOnlySet<uint>, bool, IProgress<DeviceImportProgress>, CancellationToken, Task<DeviceImportResult>> save,
    string serverName = "Current server") : ObservableObject
{
    private readonly TaskCompletionSource<DeviceImportResult?> completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public Task<DeviceImportResult?> Completion => completion.Task;
    public string SourceDescription => "Shared devices for " + serverName;
    [ObservableProperty] private IReadOnlyList<DeviceImportRowViewModel> items = [];
    [ObservableProperty] private string search = "";
    [ObservableProperty] private string message = "Load shared devices for the connected server. Nothing is imported until you confirm.";
    [ObservableProperty] private bool acknowledgeUnknownWipe;
    [ObservableProperty] private bool isBusy;
    [ObservableProperty] private double progressPercent;
    [ObservableProperty] private string checkingDevice = "";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsIndeterminate))]
    private int progressTotal;
    public bool IsIndeterminate => ProgressTotal == 0;
    private DeviceImportPreview? preview;
    private DeviceImportResult? result;
    public IReadOnlyList<DeviceImportRowViewModel> VisibleItems => Items.Where(row => string.IsNullOrWhiteSpace(Search) ||
        (row.Name + " " + row.Owner + " " + row.Item.Device.Kind).Contains(Search.Trim(), StringComparison.OrdinalIgnoreCase)).ToArray();
    public bool HasNoMatches => Items.Count > 0 && VisibleItems.Count == 0;
    public string SelectionCaption => $"{Items.Count(row => row.IsSelected)} selected · {Items.Count} shared devices";
    public bool NeedsWipeAcknowledgment => Items.Any(row => row.IsSelected && row.Item.UnknownWipe);
    public bool CanImport => !IsBusy && preview is not null && result is null && Items.Any(row => row.IsSelected) &&
        (!NeedsWipeAcknowledgment || AcknowledgeUnknownWipe);
    public bool CanCheck => !IsBusy && preview is not null && result is null && Items.Any(row => row.Item.State != DeviceImportState.Unsupported);
    public bool CanLoad => !IsBusy && result is null;
    partial void OnSearchChanged(string value) { OnPropertyChanged(nameof(VisibleItems)); OnPropertyChanged(nameof(HasNoMatches)); }
    partial void OnIsBusyChanged(bool value) { ProgressTotal = 0; ProgressPercent = 0; CheckingDevice = ""; NotifyState(); }
    partial void OnAcknowledgeUnknownWipeChanged(bool value) => NotifyState();

    [RelayCommand(CanExecute = nameof(CanLoad))]
    private async Task LoadAsync(CancellationToken ct)
    {
        if (!CanLoad) return;
        IsBusy = true; preview = null; Items = []; AcknowledgeUnknownWipe = false;
        try
        {
            preview = await load(ct);
            Items = preview.Candidates.Select(item => new DeviceImportRowViewModel(item)).ToArray();
            foreach (var item in Items) item.PropertyChanged += (_, args) => { if (args.PropertyName == nameof(item.IsSelected)) NotifyState(); };
            Message = Items.Count == 0 ? "No shared device snapshots were found for this team and server. " + preview.Notice : preview.Notice;
        }
        catch (Exception ex) { Message = Failure(ex); }
        finally { IsBusy = false; OnPropertyChanged(nameof(VisibleItems)); OnPropertyChanged(nameof(HasNoMatches)); NotifyState(); }
    }

    [RelayCommand(CanExecute = nameof(CanCheck))]
    private async Task CheckAsync(CancellationToken ct)
    {
        if (!CanCheck || preview is null) return;
        IsBusy = true;
        try
        {
            var checkedDevices = await check(preview, Items.Where(row => row.Item.State != DeviceImportState.Unsupported).Select(row => row.Item.Device.EntityId).ToHashSet(), Progress(), ct);
            foreach (var row in Items)
            {
                if (checkedDevices.Available.Contains(row.Item.Device.EntityId)) row.Status = row.Item.State switch
                {
                    DeviceImportState.AlreadyPresent => "Already in your saved tree · responding",
                    DeviceImportState.PreviousWipe => "Previous wipe · ID responds; cannot import",
                    _ => "Responding on the connected server"
                };
                if (checkedDevices.Unavailable.Contains(row.Item.Device.EntityId)) { row.Status = "Missing or unavailable on the connected server"; row.IsSelected = false; }
            }
            Message = $"Checked devices: {checkedDevices.Available.Count} responding; {checkedDevices.Unavailable.Count} unavailable. Saving checks the selection again.";
        }
        catch (Exception ex) { Message = Failure(ex); }
        finally { IsBusy = false; }
    }

    [RelayCommand(CanExecute = nameof(CanImport))]
    private async Task ImportAsync(CancellationToken ct)
    {
        if (!CanImport || preview is null) return;
        IsBusy = true;
        try
        {
            result = await save(preview, Selected(), AcknowledgeUnknownWipe, Progress(), ct);
            Message = $"Imported {result.AddedDevices} devices. Close to reload your saved device tree.";
        }
        catch (Exception ex) { Message = Failure(ex); }
        finally { IsBusy = false; }
    }

    [RelayCommand] private void SelectAll() { if (!IsBusy) foreach (var row in Items.Where(row => row.CanChoose)) row.IsSelected = true; }
    [RelayCommand] private void ClearSelection() { if (!IsBusy) foreach (var row in Items) row.IsSelected = false; }
    private IReadOnlySet<uint> Selected() => Items.Where(row => row.IsSelected && row.CanChoose).Select(row => row.Item.Device.EntityId).ToHashSet();
    private IProgress<DeviceImportProgress> Progress() => new Progress<DeviceImportProgress>(value =>
    {
        if (!IsBusy) return;
        ProgressTotal = value.Total;
        ProgressPercent = value.Total == 0 ? 0 : value.Completed * 100d / value.Total;
        CheckingDevice = $"{value.Completed} / {value.Total}: {value.Device}";
    });
    [RelayCommand] private async Task CloseAsync() { await StopAsync(); completion.TrySetResult(result); }
    public async Task StopAsync()
    {
        LoadCommand.Cancel(); CheckCommand.Cancel(); ImportCommand.Cancel();
        foreach (var task in new[] { LoadCommand.ExecutionTask, CheckCommand.ExecutionTask, ImportCommand.ExecutionTask })
            if (task is not null) await task;
    }

    private void NotifyState()
    {
        OnPropertyChanged(nameof(SelectionCaption)); OnPropertyChanged(nameof(NeedsWipeAcknowledgment));
        OnPropertyChanged(nameof(CanLoad)); OnPropertyChanged(nameof(CanCheck)); OnPropertyChanged(nameof(CanImport));
        LoadCommand.NotifyCanExecuteChanged(); CheckCommand.NotifyCanExecuteChanged(); ImportCommand.NotifyCanExecuteChanged();
    }
    private static string Failure(Exception ex) => ex switch
    {
        OperationCanceledException => "Device import was canceled or the connection changed. Load again before retrying.",
        DeviceImportException { Failure: DeviceImportFailure.ChangedPreview } => "Shared devices, team membership or the wipe changed. Load the preview again.",
        DeviceImportException { Failure: DeviceImportFailure.UnverifiedDevices } => "The current wipe or selected devices could not be verified. Check the server, then load again.",
        DeviceImportException { Failure: DeviceImportFailure.Disconnected } => "The connection changed. Reconnect and load the preview again.",
        _ => "Unable to import these devices safely. Refresh the server list and check source files, permissions and the connection."
    };
}
