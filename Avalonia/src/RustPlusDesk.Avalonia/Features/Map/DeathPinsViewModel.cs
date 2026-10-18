using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>
/// Reference death markers (ServerProfile.DeathMarkers, ChkDeathMarkers/MapShowDeathTags, RedrawDeathPins, the Layers
/// caps and wipe, DeathMarkerSettingsDialog): every team member seen dying (alive → dead between two team reads) gets a
/// pin at their position (else their player marker), the oldest pins beyond the per-player caps are dropped, and the
/// selected profile keeps them in profiles.json. Pins show on the map while the "Death markers" switch is on.
/// </summary>
public sealed partial class DeathPinsViewModel : ObservableObject
{
    private readonly IDeathMarkerStore store;
    private readonly IAppSettingsStore? settings;
    private readonly MapViewModel map;
    private readonly Func<(string Path, string ProfileKey)?> profile;
    private readonly Func<string?> ownSteamId;
    private readonly PresenceTracker presence = new();
    private IReadOnlyList<DeathMarker> markers = [];
    private int loadVersion;

    public DeathPinsViewModel(IDeathMarkerStore store, IAppSettingsStore? settings, TeamViewModel team, MapViewModel map,
        Func<(string Path, string ProfileKey)?> profile, Func<string?> ownSteamId)
    {
        this.store = store;
        this.settings = settings;
        this.map = map;
        this.profile = profile;
        this.ownSteamId = ownSteamId;
        show = settings?.Get(SettingKeys.MapShowDeathTags, false) ?? false;
        maxSelf = DeathMarkerRules.ClampCap(settings?.Get(SettingKeys.MaxSelfDeathMarkers, DeathMarkerRules.DefaultCap) ?? DeathMarkerRules.DefaultCap);
        maxTeam = DeathMarkerRules.ClampCap(settings?.Get(SettingKeys.MaxTeamDeathMarkers, DeathMarkerRules.DefaultCap) ?? DeathMarkerRules.DefaultCap);
        team.SnapshotRead += snapshot => _ = ObserveAsync(snapshot.Members);
        team.Cleared += presence.Reset;
        map.PropertyChanged += (_, args) => { if (args.PropertyName == nameof(MapViewModel.AbbreviateNames)) OnPropertyChanged(nameof(Pins)); };
    }

    public Action<string>? Log { get; set; }
    /// <summary>Clock for new markers (reference DateTime.Now); replaced by tests.</summary>
    public Func<DateTime> Now { get; set; } = () => DateTime.Now;

    /// <summary>Reference MapShowDeathTags (default off).</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowWipeOverlay))]
    private bool show;
    partial void OnShowChanged(bool value) => settings?.Set(SettingKeys.MapShowDeathTags, value);

    /// <summary>Reference MaxSelfDeathMarkers (1–50, default 3): pins kept for the own player.</summary>
    [ObservableProperty] private int maxSelf;
    partial void OnMaxSelfChanged(int value)
    {
        int clamped = DeathMarkerRules.ClampCap(value);
        if (clamped != value) { MaxSelf = clamped; return; }
        settings?.Set(SettingKeys.MaxSelfDeathMarkers, value);
    }

    /// <summary>Reference MaxTeamDeathMarkers (1–50, default 3): pins kept per teammate.</summary>
    [ObservableProperty] private int maxTeam;
    partial void OnMaxTeamChanged(int value)
    {
        int clamped = DeathMarkerRules.ClampCap(value);
        if (clamped != value) { MaxTeam = clamped; return; }
        settings?.Set(SettingKeys.MaxTeamDeathMarkers, value);
    }

    /// <summary>The Layers wipe asks first (reference ConfirmAndClearDeathMarkers).</summary>
    [ObservableProperty] private bool isWipeConfirming;

    public IReadOnlyList<DeathMarker> Markers => markers;
    /// <summary>The island map's pins, numbered and labelled as the reference does; names follow streamer mode.</summary>
    public IReadOnlyList<DeathPin> Pins => DeathMarkerRules.Pins(markers.Select(marker => map.AbbreviateNames
        ? marker with { OriginalName = StreamerMode.Name(marker.OriginalName, true), CustomName = string.IsNullOrEmpty(marker.CustomName) ? "" : StreamerMode.Name(marker.CustomName, true) }
        : marker));
    public bool HasPins => Pins.Count > 0;
    /// <summary>Reference WipeDeathMarkersOverlay: shown on the map while pins are shown and there are any.</summary>
    public bool ShowWipeOverlay => Show && HasPins;
    public string CountText => Pins.Count == 1 ? "1 death marker" : Pins.Count.ToString(CultureInfo.InvariantCulture) + " death markers";

    /// <summary>Loads the selected profile's markers (server selection).</summary>
    public async Task ReloadAsync(CancellationToken ct = default)
    {
        int version = ++loadVersion;
        IReadOnlyList<DeathMarker> loaded = [];
        if (profile() is { } target)
        {
            try { loaded = await store.LoadAsync(target.Path, target.ProfileKey, ct); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException) { Log?.Invoke("[deaths] Death markers could not be read."); }
        }
        if (version == loadVersion) Set(loaded);
    }

    /// <summary>Reference death capture from team reads (alive → dead after a first sighting).</summary>
    public async Task ObserveAsync(IReadOnlyList<ServerTeamMember> members)
    {
        foreach (var change in presence.Observe(members))
        {
            if (change.WasDead || !change.IsDead || !ulong.TryParse(change.Member.SteamId, NumberStyles.None, CultureInfo.InvariantCulture, out ulong steamId)) continue;
            if (Position(change.Member, steamId) is not { } at) continue;
            var marker = new DeathMarker(Guid.NewGuid(), steamId, change.Member.Name, "", at.X, at.Y, Now());
            ulong? self = OwnId();
            int keepSelf = MaxSelf, keepTeam = MaxTeam;
            await ChangeAsync(list => DeathMarkerRules.Add(list, marker, self, keepSelf, keepTeam));
        }
    }

    /// <summary>Reference BtnWipeDeathMarkers_Click (map overlay, no question) and the confirmed Layers / settings wipe.</summary>
    [RelayCommand]
    private Task WipeAsync()
    {
        IsWipeConfirming = false;
        return ChangeAsync(_ => []);
    }

    /// <summary>Layers "Wipe death markers": the first click asks, the second wipes.</summary>
    [RelayCommand]
    private Task ConfirmWipeAsync()
    {
        if (!IsWipeConfirming) { IsWipeConfirming = true; return Task.CompletedTask; }
        return WipeAsync();
    }

    [RelayCommand] private void CancelWipe() => IsWipeConfirming = false;

    /// <summary>Reference DeathMarkerSettingsDialog Save: trims every player to the current caps now.</summary>
    [RelayCommand]
    private Task ApplyCapsAsync()
    {
        ulong? self = OwnId();
        int keepSelf = MaxSelf, keepTeam = MaxTeam;
        return ChangeAsync(list => DeathMarkerRules.Trim(list, self, keepSelf, keepTeam));
    }

    /// <summary>Reference RenameDeathMarker_Click: an empty name restores the player's name.</summary>
    public Task RenameAsync(Guid id, string? name)
    {
        string custom = string.IsNullOrWhiteSpace(name) ? "" : name;
        return ChangeAsync(list => list.Select(marker => marker.Id == id ? marker with { CustomName = custom } : marker).ToArray());
    }

    /// <summary>Reference DeleteDeathMarker_Click.</summary>
    public Task DeleteAsync(Guid id) => ChangeAsync(list => list.Where(marker => marker.Id != id).ToArray());

    private async Task ChangeAsync(Func<IReadOnlyList<DeathMarker>, IReadOnlyList<DeathMarker>> change)
    {
        if (profile() is not { } target) return;
        try
        {
            var saved = await store.UpdateAsync(target.Path, target.ProfileKey, change, CancellationToken.None);
            ++loadVersion;
            Set(saved);
        }
        catch (ProfileChangedException) { Log?.Invoke("[deaths] The server profile changed; death markers were not saved."); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or ArgumentException)
        { Log?.Invoke("[deaths] Death markers could not be saved."); }
    }

    private (double X, double Y)? Position(ServerTeamMember member, ulong steamId)
    {
        if (member.X is { } x && member.Y is { } y) return (x, y);
        // Reference TryResolvePosFromDynMarkers: the player's live marker.
        return map.Markers.FirstOrDefault(marker => marker.Kind == MapMarkerKind.Player && marker.SteamId == steamId) is { } live ? (live.X, live.Y) : null;
    }

    private ulong? OwnId() => ulong.TryParse(ownSteamId(), NumberStyles.None, CultureInfo.InvariantCulture, out ulong id) && id != 0 ? id : null;

    private void Set(IReadOnlyList<DeathMarker> value)
    {
        markers = value;
        foreach (string name in new[] { nameof(Markers), nameof(Pins), nameof(HasPins), nameof(ShowWipeOverlay), nameof(CountText) }) OnPropertyChanged(name);
    }
}
