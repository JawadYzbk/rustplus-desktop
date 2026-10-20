using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Devices;

/// <summary>One condition a rule can use, as the editor's combo box shows it.</summary>
public sealed record AutomationChoice(string Label, string Tag)
{
    public override string ToString() => Label;
}

/// <summary>
/// Reference DeviceAutomationRule: one rule that puts a paired switch into a state when its condition holds.
/// </summary>
public sealed partial class AutomationRuleViewModel : ObservableObject
{
    public static IReadOnlyList<AutomationChoice> ConditionTypes { get; } =
        [new("Player proximity", "PlayerProximity"), new("Game time", "GameTime")];

    public static IReadOnlyList<AutomationChoice> PlayerMatchModes { get; } =
    [
        new("Any player online", "AnyOnline"),
        new("All players online", "AllOnline"),
        new("Any player offline", "AnyOffline"),
        new("All players offline", "AllOffline"),
        new("A specific player", "Specific"),
        new("A specific player offline", "SpecificOffline")
    ];

    public static IReadOnlyList<AutomationChoice> UnmatchedChoices { get; } =
        [new("On", "On"), new("Off", "Off"), new("Leave alone", "None")];

    [ObservableProperty] private string name = "New Automation";
    [ObservableProperty] private bool isEnabled;
    [ObservableProperty] private AutomationChoice conditionType = ConditionTypes[0];
    [ObservableProperty] private AutomationChoice playerMatchMode = PlayerMatchModes[0];
    [ObservableProperty] private string specificPlayerSteamId = "";
    [ObservableProperty] private AutomationChoice? locationDevice;
    [ObservableProperty] private AutomationChoice? targetDevice;
    [ObservableProperty] private string distanceMeters = "250";
    [ObservableProperty] private string startTime = "20:00";
    [ObservableProperty] private string endTime = "08:00";
    [ObservableProperty] private bool matchedState;
    [ObservableProperty] private AutomationChoice unmatchedChoice = UnmatchedChoices[2];
    [ObservableProperty] private string? lastAppliedText;

    public string Id { get; init; } = Guid.NewGuid().ToString();

    /// <summary>Whether the rule names one player, so the Steam ID box belongs on screen.</summary>
    public bool IsSpecificMode => PlayerMatchMode.Tag.StartsWith("Specific", StringComparison.Ordinal);

    partial void OnPlayerMatchModeChanged(AutomationChoice value) => OnPropertyChanged(nameof(IsSpecificMode));

    /// <summary>The rule as the core model reads it.</summary>
    public DeviceAutomationRule ToRule() => new()
    {
        Id = Id,
        Name = string.IsNullOrWhiteSpace(Name) ? "New Automation" : Name.Trim(),
        IsEnabled = IsEnabled,
        ConditionType = ConditionType.Tag,
        PlayerMatchMode = PlayerMatchMode.Tag,
        SpecificPlayerSteamId = ulong.TryParse(SpecificPlayerSteamId.Trim(), out ulong steamId) ? steamId : 0,
        LocationEntityId = LocationDevice is null ? 0 : LocationDevice.Tag switch
        {
            string text when uint.TryParse(text, out uint id) => id,
            _ => 0
        },
        DistanceMeters = double.TryParse(DistanceMeters, NumberStyles.Float, CultureInfo.InvariantCulture, out double distance)
            ? Math.Max(1, distance) : 250,
        StartTime = string.IsNullOrWhiteSpace(StartTime) ? "20:00" : StartTime.Trim(),
        EndTime = string.IsNullOrWhiteSpace(EndTime) ? "08:00" : EndTime.Trim(),
        TargetEntityId = TargetDevice is null ? 0 : TargetDevice.Tag switch
        {
            string text when uint.TryParse(text, out uint id) => id,
            _ => 0
        },
        MatchedState = MatchedState,
        UnmatchedState = UnmatchedChoice.Tag switch { "On" => true, "Off" => false, _ => null }
    };

    /// <summary>Reads a saved rule back into the editor.</summary>
    public static AutomationRuleViewModel FromRule(DeviceAutomationRule rule, IReadOnlyList<AutomationChoice> devices)
    {
        var view = new AutomationRuleViewModel
        {
            Id = rule.Id,
            Name = rule.Name,
            IsEnabled = rule.IsEnabled,
            ConditionType = ConditionTypes.FirstOrDefault(choice => choice.Tag == rule.ConditionType) ?? ConditionTypes[0],
            PlayerMatchMode = PlayerMatchModes.FirstOrDefault(choice => choice.Tag == rule.PlayerMatchMode) ?? PlayerMatchModes[0],
            SpecificPlayerSteamId = rule.SpecificPlayerSteamId == 0 ? "" : rule.SpecificPlayerSteamId.ToString(CultureInfo.InvariantCulture),
            DistanceMeters = rule.DistanceMeters.ToString("0.##", CultureInfo.InvariantCulture),
            StartTime = rule.StartTime,
            EndTime = rule.EndTime,
            MatchedState = rule.MatchedState,
            UnmatchedChoice = UnmatchedChoices.FirstOrDefault(choice => choice.Tag == (rule.UnmatchedState switch { true => "On", false => "Off", _ => "None" })) ?? UnmatchedChoices[2]
        };
        view.LocationDevice = devices.FirstOrDefault(choice => choice.Tag == rule.LocationEntityId.ToString(CultureInfo.InvariantCulture));
        view.TargetDevice = devices.FirstOrDefault(choice => choice.Tag == rule.TargetEntityId.ToString(CultureInfo.InvariantCulture));
        return view;
    }
}

/// <summary>
/// Reference DeviceAutomationOverlay: the rules of the selected profile, the switch that turns them on and the
/// editor for one rule. The rules live in the profile, so they are written through the backed-up profile writer.
/// </summary>
public sealed partial class DeviceAutomationSettingsViewModel(IDeviceAutomationStore store,
    Func<(string Path, ServerSummary Server)?> selected, Func<IReadOnlyList<PairedDevice>>? devices = null,
    Action<DeviceAutomationSettings>? applied = null) : ObservableObject
{
    public const string NoServerStatus = "Select a server to configure its device automation.";
    public const string SavedStatus = "Device automation saved.";
    public const string ChangedStatus = "The profiles file changed. Refresh the server list and try again.";
    public const string SaveFailedStatus = "Unable to save. Check file permissions and available disk space.";


    [ObservableProperty] private bool enabled;    [ObservableProperty] private string status = NoServerStatus;
    [ObservableProperty] private bool isError;

    public ObservableCollection<AutomationRuleViewModel> Rules { get; } = [];

    /// <summary>The paired devices the rules can point at.</summary>
    public IReadOnlyList<AutomationChoice> DeviceChoices { get; private set; } = [];

    /// <summary>Reference LoadAsync: the selected profile's rules.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        IsError = false;
        DeviceChoices = Choices();
        if (selected() is not { } target)
        {
            Enabled = false;
            Rules.Clear();
            Status = NoServerStatus;
            return;
        }

        DeviceAutomationSettings settings;
        try { settings = await store.GetAsync(target.Path, target.Server.ProfileKey, ct); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            settings = DeviceAutomationSettings.Defaults;
        }

        Enabled = settings.Enabled;
        Rules.Clear();
        foreach (var rule in settings.Rules) Rules.Add(AutomationRuleViewModel.FromRule(rule, DeviceChoices));
        Status = SavedStatus;
        applied?.Invoke(settings);
    }

    /// <summary>Reference AddAutomation: a new rule starts in the editor, switched off.</summary>
    [RelayCommand]
    private void AddRule()
    {
        Rules.Add(new AutomationRuleViewModel());
        Status = "Add the rule's condition and target, then apply.";
    }

    [RelayCommand]
    private void RemoveRule(AutomationRuleViewModel? rule)
    {
        if (rule is null) return;
        Rules.Remove(rule);
    }

    /// <summary>Reference Save: the rules and the switch are written to the profile.</summary>
    [RelayCommand]
    private async Task ApplyAsync(CancellationToken ct)
    {
        if (selected() is null)
        {
            (Status, IsError) = (NoServerStatus, true);
            return;
        }

        var settings = new DeviceAutomationSettings(Enabled, Rules.Select(rule => rule.ToRule()).ToList());
        Enabled = settings.Enabled;
        await PersistAsync(settings, ct);
    }

    private async Task PersistAsync(DeviceAutomationSettings settings, CancellationToken ct)
    {
        if (selected() is not { } target)
        {
            (Status, IsError) = (NoServerStatus, true);
            return;
        }
        try { await store.SaveAsync(target.Path, target.Server.ProfileKey, settings, ct); }
        catch (ProfileChangedException) { (Status, IsError) = (ChangedStatus, true); return; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            (Status, IsError) = (SaveFailedStatus, true);
            return;
        }
        applied?.Invoke(settings);
        (Status, IsError) = (SavedStatus, false);
    }

    /// <summary>The paired devices, as the rule editor's pickers show them.</summary>
    private IReadOnlyList<AutomationChoice> Choices()
    {
        var choices = new List<AutomationChoice>();
        foreach (var device in devices?.Invoke() ?? [])
        {
            if (device.Kind == DeviceKind.Group) continue;
            choices.Add(new AutomationChoice($"{device.Name} (#{device.EntityId})", device.EntityId.ToString(CultureInfo.InvariantCulture)));
        }
        return choices;
    }
}
