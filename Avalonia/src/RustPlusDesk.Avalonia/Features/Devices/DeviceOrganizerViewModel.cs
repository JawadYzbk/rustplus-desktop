using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Devices;

public sealed record DeviceGroupChoice(string Name, string? DeviceKey);

/// <summary>Owns revision-protected local organization and awaits writes before closing.</summary>
public sealed partial class DeviceOrganizerViewModel : ObservableObject
{
    private readonly string path;
    private readonly IDeviceProfileEditor editor;
    private readonly Func<Task<string?>>? pickSound;
    private ServerSummary current;
    private bool changed;
    private bool closing;
    private readonly TaskCompletionSource<ServerSummary?> completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public Task<ServerSummary?> Completion => completion.Task;
    public IReadOnlyList<PairedDevice> Devices => current.Devices;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RenameCommand))]
    [NotifyCanExecuteChangedFor(nameof(SaveIconCommand))]
    [NotifyCanExecuteChangedFor(nameof(ResetIconCommand))]
    [NotifyCanExecuteChangedFor(nameof(SaveAlarmTitleCommand))]
    [NotifyPropertyChangedFor(nameof(IsAlarm))]
    [NotifyCanExecuteChangedFor(nameof(GroupCommand))]
    [NotifyCanExecuteChangedFor(nameof(MoveCommand))]
    [NotifyCanExecuteChangedFor(nameof(MoveUpCommand))]
    [NotifyCanExecuteChangedFor(nameof(MoveDownCommand))]
    [NotifyCanExecuteChangedFor(nameof(DeleteCommand))]
    private PairedDevice? selected;
    [ObservableProperty] private string alias = "";
    private readonly IReadOnlyList<DeviceIcon> icons;
    [ObservableProperty] private string iconSearch = "";
    [ObservableProperty] private bool isIconPickerOpen;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SaveIconCommand))]
    private DeviceIcon? selectedIcon;
    public bool HasIconCatalog => icons.Count > 0;
    public IEnumerable<DeviceIcon> IconChoices => icons.Where(icon => icon.Name.Contains(IconSearch, StringComparison.OrdinalIgnoreCase)
        || icon.ShortName.Contains(IconSearch, StringComparison.OrdinalIgnoreCase) || icon.Id.ToString(CultureInfo.InvariantCulture).Contains(IconSearch, StringComparison.Ordinal));
    partial void OnIconSearchChanged(string value) => OnPropertyChanged(nameof(IconChoices));
    private bool CanSaveIcon => CanEdit && SelectedIcon is not null;
    [ObservableProperty] private string alarmTitle = "";
    public bool IsAlarm => Selected?.Kind == DeviceKind.Alarm;
    private bool CanEditAlarm => CanEdit && IsAlarm;
    /// <summary>Reference SmartDevice.AudioFilePath: the sound this alarm makes, when it has one of its own.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAlarmAudio))]
    private string alarmAudioPath = "";

    /// <summary>Whether this alarm has a sound file of its own.</summary>
    public bool HasAlarmAudio => !string.IsNullOrWhiteSpace(AlarmAudioPath);
    [ObservableProperty] private string groupName = "Group";
    [ObservableProperty] private IReadOnlyList<DeviceGroupChoice> destinations = [];
    [ObservableProperty] private DeviceGroupChoice? destination;
    [ObservableProperty] private string message = "Each action saves locally. Enabled device sync resumes after you close this editor.";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RenameCommand))]
    [NotifyCanExecuteChangedFor(nameof(SaveIconCommand))]
    [NotifyCanExecuteChangedFor(nameof(ResetIconCommand))]
    [NotifyCanExecuteChangedFor(nameof(SaveAlarmTitleCommand))]
    [NotifyCanExecuteChangedFor(nameof(GroupCommand))]
    [NotifyCanExecuteChangedFor(nameof(MoveCommand))]
    [NotifyCanExecuteChangedFor(nameof(MoveUpCommand))]
    [NotifyCanExecuteChangedFor(nameof(MoveDownCommand))]
    [NotifyCanExecuteChangedFor(nameof(DeleteCommand))]
    private bool isBusy;
    /// <summary>Deletion is irreversible, so the first press only asks for confirmation.</summary>
    [ObservableProperty] private bool isDeleteConfirming;
    public string DeleteConfirmation => Selected is not { } item ? "" : item.Kind == DeviceKind.Group
        ? $"Delete \"{item.Name}\" and the {Leaves(item)} device(s) inside it? Owned oil-rig rules and chat commands are removed too."
        : $"Delete \"{item.Name}\"? Its oil-rig rules and chat command are removed too.";
    private static int Leaves(PairedDevice item) => item.DescendantsAndSelf().Count(device => device.Kind != DeviceKind.Group);
    private bool needsRefresh;
    private bool CanEdit => !closing && !IsBusy && !needsRefresh && Selected is not null;
    private bool CanMoveUp => CanEdit && Position > 0;
    private bool CanMoveDown => CanEdit && Position < Siblings.Count - 1;
    private int Position => Selected is null ? -1 : int.Parse(Selected.DeviceKey.Split('/')[^1], CultureInfo.InvariantCulture);
    private string? ParentKey => Selected?.DeviceKey.LastIndexOf('/') is int slash && slash >= 0 ? Selected.DeviceKey[..slash] : null;
    private IReadOnlyList<PairedDevice> Siblings => ParentKey is { } key
        ? Devices.SelectMany(item => item.DescendantsAndSelf()).First(item => item.DeviceKey == key).Children : Devices;

    public DeviceOrganizerViewModel(string path, ServerSummary original, IDeviceProfileEditor editor,
        IReadOnlyList<DeviceIcon>? icons = null, Func<Task<string?>>? pickSound = null)
    {
        this.path = path; current = original; this.editor = editor;
        this.icons = icons ?? [];
        this.pickSound = pickSound;
        Selected = Devices.FirstOrDefault();
    }

    partial void OnSelectedChanged(PairedDevice? value)
    {
        IsDeleteConfirming = false;
        OnPropertyChanged(nameof(DeleteConfirmation));
        Alias = value?.Alias ?? "";
        AlarmTitle = value?.InGameAlarmTitle ?? "";
        AlarmAudioPath = value?.AudioFilePath ?? "";
        IconSearch = "";
        SelectedIcon = icons.FirstOrDefault(icon => icon.Id == value?.CustomIconId || icon.ShortName == value?.CustomIconShortName);
        var choices = new List<DeviceGroupChoice> { new("All devices", null) };
        void Add(IEnumerable<PairedDevice> items, string prefix)
        {
            foreach (var item in items)
            {
                if (value is not null && (item.DeviceKey == value.DeviceKey || item.DeviceKey.StartsWith(value.DeviceKey + "/", StringComparison.Ordinal))) continue;
                string name = prefix + item.Name;
                if (item.Kind == DeviceKind.Group) choices.Add(new(name, item.DeviceKey));
                Add(item.Children, name + " / ");
            }
        }
        Add(Devices, "");
        Destinations = choices;
        Destination = choices.FirstOrDefault(item => item.DeviceKey == ParentKey) ?? choices[0];
    }

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private Task RenameAsync(CancellationToken ct) => SaveAsync((server, key) => editor.RenameAsync(path, server, key, Alias, ct));
    [RelayCommand(CanExecute = nameof(CanSaveIcon))]
    private Task SaveIconAsync(CancellationToken ct) => SaveAsync((server, key) => editor.SaveIconAsync(path, server, key, SelectedIcon!.Id, SelectedIcon.ShortName, ct));
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private Task ResetIconAsync(CancellationToken ct) => SaveAsync((server, key) => editor.SaveIconAsync(path, server, key, null, null, ct));
    [RelayCommand(CanExecute = nameof(CanEditAlarm))]
    private Task SaveAlarmTitleAsync(CancellationToken ct) => SaveAsync((server, key) => editor.SaveAlarmTitleAsync(path, server, key, AlarmTitle, ct));

    /// <summary>Reference OnSelectAlarmAudioFileClick: the user chooses the sound this alarm makes.</summary>
    [RelayCommand(CanExecute = nameof(CanEditAlarm))]
    private async Task PickAlarmAudioAsync()
    {
        if (pickSound is not { } pick) return;
        string? file = await pick();
        if (string.IsNullOrWhiteSpace(file)) return;
        AlarmAudioPath = file;
        await SaveAudioAsync(CancellationToken.None);
    }

    /// <summary>Reference OnResetAlarmAudioFileClick: the sound goes back to the one the app ships.</summary>
    [RelayCommand(CanExecute = nameof(CanEditAlarm))]
    private async Task ResetAlarmAudioAsync()
    {
        AlarmAudioPath = "";
        await SaveAudioAsync(CancellationToken.None);
    }

    private Task SaveAudioAsync(CancellationToken ct) =>
        SaveAsync((server, key) => editor.SaveAlarmAudioAsync(path, server, key, AlarmAudioPath, ct));
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private Task GroupAsync(CancellationToken ct) => SaveAsync((server, key) => editor.GroupAsync(path, server, key, GroupName, ct));
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private Task MoveAsync(CancellationToken ct) => SaveAsync((server, key) => editor.MoveAsync(path, server, key, Destination?.DeviceKey, null, ct));
    [RelayCommand(CanExecute = nameof(CanMoveUp))]
    private Task MoveUpAsync(CancellationToken ct) => ReorderAsync(-1, ct);
    [RelayCommand(CanExecute = nameof(CanMoveDown))]
    private Task MoveDownAsync(CancellationToken ct) => ReorderAsync(1, ct);

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task DeleteAsync(CancellationToken ct)
    {
        if (!IsDeleteConfirming) { IsDeleteConfirming = true; return; }
        IsDeleteConfirming = false;
        if (!CanEdit) return;
        string? parent = ParentKey;
        int position = Position;
        IsBusy = true;
        try
        {
            var removing = Selected!.DescendantsAndSelf().Where(device => device.Kind != DeviceKind.Group).Select(device => (long)device.EntityId).ToHashSet();
            var result = await editor.DeleteAsync(path, current, Selected!.DeviceKey, ct);
            changed |= result.Server.SourceRevision != current.SourceRevision;
            // An entity paired twice keeps its hotkey while another copy remains.
            removing.ExceptWith(result.Server.Devices.SelectMany(device => device.DescendantsAndSelf()).Select(device => (long)device.EntityId));
            if (removing.Count > 0) DevicesRemoved?.Invoke(removing);
            current = result.Server;
            OnPropertyChanged(nameof(Devices));
            var all = Devices.SelectMany(item => item.DescendantsAndSelf()).ToArray();
            string prefix = parent is null ? "" : parent + "/";
            Selected = all.FirstOrDefault(item => item.DeviceKey == prefix + position.ToString(CultureInfo.InvariantCulture))
                ?? all.FirstOrDefault(item => item.DeviceKey == prefix + (position - 1).ToString(CultureInfo.InvariantCulture))
                ?? all.FirstOrDefault(item => item.DeviceKey == parent) ?? Devices.FirstOrDefault();
            Message = $"Deleted {result.RemovedDevices} device(s)" + (result.RemovedRules > 0 ? $", {result.RemovedRules} oil-rig rule(s)" : "")
                + (result.RemovedCommands > 0 ? $", {result.RemovedCommands} chat command(s)" : "") + ". Enabled device sync uploads the change after you close this editor.";
        }
        catch (OperationCanceledException) { Message = "Delete canceled."; }
        catch (ProfileChangedException) { needsRefresh = true; Message = "The profiles file changed. Close and refresh before organizing again."; }
        catch (Exception) { Message = "Unable to delete. Check file permissions and available disk space, then close and refresh."; }
        finally { IsBusy = false; }
    }

    [RelayCommand]
    private void CancelDelete() => IsDeleteConfirming = false;

    /// <summary>Entity IDs no longer paired after a delete (hotkey cleanup).</summary>
    public Action<IReadOnlyCollection<long>>? DevicesRemoved { get; set; }

    private Task ReorderAsync(int delta, CancellationToken ct)
        => Position + delta < 0 || Position + delta >= Siblings.Count ? Task.CompletedTask
            : SaveAsync((server, key) => editor.MoveAsync(path, server, key, ParentKey, Position + delta, ct));

    private async Task SaveAsync(Func<ServerSummary, string, Task<DeviceEditResult>> save)
    {
        if (!CanEdit) return;
        IsBusy = true;
        try
        {
            var result = await save(current, Selected!.DeviceKey);
            changed |= result.Server.SourceRevision != current.SourceRevision;
            current = result.Server;
            OnPropertyChanged(nameof(Devices));
            Selected = Devices.SelectMany(item => item.DescendantsAndSelf()).First(item => item.DeviceKey == result.DeviceKey);
            Message = "Saved locally. Cloud synchronization is not available in this preview.";
        }
        catch (OperationCanceledException) { Message = "Save canceled."; }
        catch (ProfileChangedException) { needsRefresh = true; Message = "The profiles file changed. Close and refresh before organizing again."; }
        catch (ArgumentException) { Message = "Enter a group name, choose a valid destination or select an icon from the catalog."; }
        catch (Exception) { Message = "Unable to save. Check file permissions and available disk space, then close and refresh."; }
        finally { IsBusy = false; }
    }

    [RelayCommand]
    private async Task CloseAsync()
    {
        closing = true;
        await StopAsync();
        completion.TrySetResult(changed ? current : null);
    }

    public async Task StopAsync()
    {
        foreach (var command in Commands) command.Cancel();
        foreach (var command in Commands) if (command.ExecutionTask is { } task) await task;
    }

    private IAsyncRelayCommand[] Commands => [RenameCommand, SaveIconCommand, ResetIconCommand, SaveAlarmTitleCommand,
        PickAlarmAudioCommand, ResetAlarmAudioCommand, GroupCommand, MoveCommand, MoveUpCommand, MoveDownCommand, DeleteCommand];
}
