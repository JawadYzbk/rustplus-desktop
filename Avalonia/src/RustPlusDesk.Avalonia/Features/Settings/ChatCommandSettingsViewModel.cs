using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference "Chat Commands" page (ChatCommandsOverlay): the enable switch, the prefix, the answer delay and the
/// per-server command names. The reference saved the page when the settings closed and its profile setters
/// normalised every name; the card saves with Apply and writes back what was stored, so what the engine matches is
/// what the card shows. The enable switch applies immediately, as the reference's checkbox did.
/// </summary>
public sealed partial class ChatCommandSettingsViewModel(IChatCommandSettingsStore store,
    Func<(string Path, ServerSummary Server)?> selected, Action<ChatCommandSettings>? applied = null) : ObservableObject
{
    public const string NoServerStatus = "Select a server to configure its chat commands.";
    public const string SavedStatus = "Chat commands saved.";
    public const string InvalidPrefixStatus = "The prefix must be one of ! . , or \\.";
    public const string ChangedStatus = "The profiles file changed. Refresh the server list and try again.";
    public const string SaveFailedStatus = "Unable to save. Check file permissions and available disk space.";

    private ChatCommandSettings loaded = ChatCommandSettings.Defaults;
    private bool loading;

    [ObservableProperty] private bool enabled;
    [ObservableProperty] private string prefix = "!";
    /// <summary>Written as text so a half-typed value does not rewrite the stored delay.</summary>
    [ObservableProperty] private string responseDelay = "0.5";
    [ObservableProperty] private string pop = "";
    [ObservableProperty] private string list = "";
    [ObservableProperty] private string time = "";
    [ObservableProperty] private string promote = "";
    [ObservableProperty] private string deepSea = "";
    [ObservableProperty] private string cargo = "";
    [ObservableProperty] private string heli = "";
    [ObservableProperty] private string vendor = "";
    [ObservableProperty] private string afk = "";
    [ObservableProperty] private string oilRig = "";
    [ObservableProperty] private string upkeepDetail = "";
    [ObservableProperty] private string baseCodes = "";
    [ObservableProperty] private string ai = "";
    [ObservableProperty] private string customTimer = "";
    [ObservableProperty] private bool clanAllowBaseCodes;
    [ObservableProperty] private bool clanAllowSwitches;
    [ObservableProperty] private string status = NoServerStatus;
    [ObservableProperty] private bool isError;

    /// <summary>Reference Base Codes editor: five labelled four-digit codes, padded with spare rows.</summary>
    public ObservableCollection<BaseCodeRowViewModel> BaseCodeRows { get; } = [];

    /// <summary>Reference settings load: the selected profile's chat-command settings.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        IsError = false;
        if (selected() is not { } target)
        {
            loading = true;
            loaded = ChatCommandSettings.Defaults;
            Enabled = loaded.Enabled;
            Prefix = loaded.EffectivePrefix;
            ResponseDelay = DelayText(loaded.ResponseDelaySeconds);
            Pop = List = Time = Promote = DeepSea = Cargo = Heli = Vendor = Afk = "";
            OilRig = UpkeepDetail = BaseCodes = Ai = CustomTimer = "";
            ClanAllowBaseCodes = ClanAllowSwitches = false;
            SetBaseCodeRows(ChatCommandSettings.NoBaseCodes);
            loading = false;
            Status = NoServerStatus;
            return;
        }

        ChatCommandSettings settings;
        try { settings = await store.GetAsync(target.Path, target.Server.ProfileKey, ct); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            settings = ChatCommandSettings.Defaults;
        }

        loading = true;
        loaded = settings;
        Enabled = settings.Enabled;
        Prefix = settings.EffectivePrefix;
        ResponseDelay = DelayText(settings.ResponseDelaySeconds);
        Pop = settings.Pop;
        List = settings.List;
        Time = settings.Time;
        Promote = settings.Promote;
        DeepSea = settings.DeepSea;
        Cargo = settings.Cargo;
        Heli = settings.Heli;
        Vendor = settings.Vendor;
        Afk = settings.Afk;
        OilRig = settings.OilRig;
        UpkeepDetail = settings.UpkeepDetail;
        BaseCodes = settings.BaseCodes;
        Ai = settings.Ai;
        CustomTimer = settings.CustomTimer;
        ClanAllowBaseCodes = settings.ClanCommandsAllowBaseCodes;
        ClanAllowSwitches = settings.ClanCommandsAllowSwitches;
        SetBaseCodeRows(settings.BaseCodeRows);
        loading = false;
        Status = SavedStatus;
        applied?.Invoke(settings);
    }

    /// <summary>Reference EnableChatCommands_Changed: the switch takes effect at once.</summary>
    partial void OnEnabledChanged(bool value)
    {
        if (loading) return;
        loaded = loaded with { Enabled = value };
        _ = PersistAsync(loaded);
    }

    [RelayCommand]
    private async Task ApplyAsync(CancellationToken ct)
    {
        if (selected() is null)
        {
            (Status, IsError) = (NoServerStatus, true);
            return;
        }

        string prefix = Prefix.Trim();
        if (!ChatCommandSettings.IsValidPrefix(prefix))
        {
            (Status, IsError) = (InvalidPrefixStatus, true);
            return;
        }

        double delay = double.TryParse(ResponseDelay, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? Math.Clamp(parsed, 0, 5)
            : loaded.ResponseDelaySeconds;
        var baseCodes = BaseCodeRows
            .Select(row => new ChatBaseCode(row.Label.Trim(), BaseCodeRowViewModel.Sanitize(row.Code)))
            .Where(code => code.Code.Length > 0)
            .Take(ChatCommandSettings.MaxBaseCodes)
            .ToList();
        var settings = loaded with
        {
            Enabled = Enabled,
            Prefix = prefix,
            ResponseDelaySeconds = delay,
            Pop = ChatCommandSettings.ValidateCommand(Pop, ChatCommandSettings.Defaults.Pop),
            List = ChatCommandSettings.ValidateCommand(List, ChatCommandSettings.Defaults.List),
            Time = ChatCommandSettings.ValidateCommand(Time, ChatCommandSettings.Defaults.Time),
            Promote = ChatCommandSettings.ValidateCommand(Promote, ChatCommandSettings.Defaults.Promote),
            DeepSea = ChatCommandSettings.ValidateCommand(DeepSea, ChatCommandSettings.Defaults.DeepSea),
            Cargo = ChatCommandSettings.ValidateCommand(Cargo, ChatCommandSettings.Defaults.Cargo),
            Heli = ChatCommandSettings.ValidateCommand(Heli, ChatCommandSettings.Defaults.Heli),
            Vendor = ChatCommandSettings.ValidateCommand(Vendor, ChatCommandSettings.Defaults.Vendor),
            Afk = ChatCommandSettings.ValidateCommand(Afk, ChatCommandSettings.Defaults.Afk),
            OilRig = ChatCommandSettings.ValidateCommand(OilRig, ChatCommandSettings.Defaults.OilRig),
            UpkeepDetail = ChatCommandSettings.ValidateCommand(UpkeepDetail, ChatCommandSettings.Defaults.UpkeepDetail),
            BaseCodes = ChatCommandSettings.ValidateCommand(BaseCodes, ChatCommandSettings.Defaults.BaseCodes),
            Ai = ChatCommandSettings.ValidateCommand(Ai, ChatCommandSettings.Defaults.Ai),
            CustomTimer = ChatCommandSettings.ValidateCommand(CustomTimer, ChatCommandSettings.Defaults.CustomTimer),
            ClanCommandsAllowBaseCodes = ClanAllowBaseCodes,
            ClanCommandsAllowSwitches = ClanAllowSwitches,
            BaseCodeRows = baseCodes.Count == 0 ? ChatCommandSettings.NoBaseCodes : baseCodes
        };

        loaded = settings;
        WriteBack(settings);
        await PersistAsync(settings, ct);
    }

    private async Task PersistAsync(ChatCommandSettings settings, CancellationToken ct = default)
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

    /// <summary>Shows what was actually stored, so a normalized name is visible before the next command runs.</summary>
    private void WriteBack(ChatCommandSettings settings)
    {
        loading = true;
        Prefix = settings.EffectivePrefix;
        Pop = settings.Pop;
        List = settings.List;
        Time = settings.Time;
        Promote = settings.Promote;
        DeepSea = settings.DeepSea;
        Cargo = settings.Cargo;
        Heli = settings.Heli;
        Vendor = settings.Vendor;
        Afk = settings.Afk;
        OilRig = settings.OilRig;
        UpkeepDetail = settings.UpkeepDetail;
        BaseCodes = settings.BaseCodes;
        Ai = settings.Ai;
        CustomTimer = settings.CustomTimer;
        ClanAllowBaseCodes = settings.ClanCommandsAllowBaseCodes;
        ClanAllowSwitches = settings.ClanCommandsAllowSwitches;
        SetBaseCodeRows(settings.BaseCodeRows);
        ResponseDelay = DelayText(settings.ResponseDelaySeconds);
        loading = false;
    }

    private void SetBaseCodeRows(IReadOnlyList<ChatBaseCode> codes)
    {
        BaseCodeRows.Clear();
        foreach (var code in codes.Take(ChatCommandSettings.MaxBaseCodes))
            BaseCodeRows.Add(new BaseCodeRowViewModel { Label = code.Label, Code = code.Code });
        while (BaseCodeRows.Count < ChatCommandSettings.MaxBaseCodes) BaseCodeRows.Add(new BaseCodeRowViewModel());
    }

    private static string DelayText(double seconds) => seconds.ToString("0.0#", CultureInfo.InvariantCulture);
}

/// <summary>
/// One Base Codes editor row. The reference ServerProfile.BaseCode setter keeps only digits and caps the length at
/// four, so a pasted code is normalised as it is typed.
/// </summary>
public sealed partial class BaseCodeRowViewModel : ObservableObject
{
    [ObservableProperty] private string code = "";
    [ObservableProperty] private string label = "";

    partial void OnCodeChanged(string value)
    {
        string digits = Sanitize(value);
        if (digits != value) Code = digits;
    }

    public static string Sanitize(string? value) =>
        new((value ?? "").Where(char.IsDigit).Take(4).ToArray());
}
