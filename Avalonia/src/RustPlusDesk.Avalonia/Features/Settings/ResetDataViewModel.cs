using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the reset data window: the components the player wants wiped from this machine.
///
/// The options are the old app's own, and each one says what it takes with it, because a wipe that takes the wrong
/// thing is not a wipe that can be undone.
/// </summary>
public sealed partial class ResetDataViewModel : ObservableObject
{
    /// <summary>Local cache and drawings: the saved maps and the lines drawn on them.</summary>
    [ObservableProperty] private bool resetCache = true;

    /// <summary>The active server connection: the app disconnects from the server.</summary>
    [ObservableProperty] private bool resetConnection = true;

    /// <summary>The custom crosshairs: the configurations are wiped.</summary>
    [ObservableProperty] private bool resetCrosshairs = true;

    /// <summary>The device pairing configuration: the registration tokens are reset.</summary>
    [ObservableProperty] private bool resetPairing = true;

    /// <summary>The saved server profiles: the paired server configurations are deleted.</summary>
    [ObservableProperty] private bool resetProfiles = true;

    /// <summary>The Steam account credentials: the player is logged out.</summary>
    [ObservableProperty] private bool resetSteam = true;

    /// <summary>What the wipe did, or why it did nothing.</summary>
    [ObservableProperty] private string status = "";

    /// <summary>Whether anything at all is chosen.</summary>
    public bool HasSelection => ResetCache || ResetConnection || ResetCrosshairs || ResetPairing || ResetProfiles || ResetSteam;

    /// <summary>Disconnects from the server, when the player asked for it.</summary>
    public Func<Task>? Disconnect { get; set; }

    /// <summary>Logs the player out, when the player asked for it.</summary>
    public Func<Task>? SignOut { get; set; }

    /// <summary>Wipes the saved maps and the drawings, when the player asked for it.</summary>
    public Action? ClearCacheAndDrawings { get; set; }

    /// <summary>Wipes the crosshair configurations, when the player asked for it.</summary>
    public Action? ClearCrosshairs { get; set; }

    /// <summary>Deletes the saved server profiles, when the player asked for it.</summary>
    public Action? ClearProfiles { get; set; }

    /// <summary>Whether the wipe happened, so the window can close itself.</summary>
    public bool DidReset { get; private set; }

    /// <summary>Called when the wipe has finished, so the window can close itself.</summary>
    public Action? Completed { get; set; }

    /// <summary>
    /// Reference BtnConfirm_Click: the chosen components are wiped, and the window says what it did.
    /// </summary>
    [RelayCommand]
    public async Task ConfirmAsync()
    {
        if (!HasSelection)
        {
            Status = Loc.Text("ResetDataWarning", "This action is permanent and cannot be undone.");
            return;
        }

        var done = new List<string>();

        try
        {
            if (ResetCache) { ClearCacheAndDrawings?.Invoke(); done.Add(Loc.Text("ResetOptionCache", "Local Cache & Drawings (Wipe saved maps/lines)")); }
            if (ResetProfiles) { ClearProfiles?.Invoke(); done.Add(Loc.Text("ResetOptionProfiles", "Saved Server Profiles (Delete paired server configurations)")); }
            if (ResetCrosshairs) { ClearCrosshairs?.Invoke(); done.Add(Loc.Text("ResetOptionCrosshairs", "Custom Crosshairs (Wipe configurations)")); }

            if (ResetConnection && Disconnect is { } disconnect) { await disconnect().ConfigureAwait(true); done.Add(Loc.Text("ResetOptionConnection", "Active Server Connection (Disconnect)")); }

            if (ResetSteam && SignOut is { } signOut) { await signOut().ConfigureAwait(true); done.Add(Loc.Text("ResetOptionSteam", "Steam Account Credentials (Log out)")); }
            if (ResetPairing && SignOut is { } pairing) { await pairing().ConfigureAwait(true); done.Add(Loc.Text("ResetOptionPairing", "Device Pairing Config (Reset registration tokens)")); }
        }
        catch (Exception ex)
        {
            Status = $"The wipe did not finish: {ex.Message}";
            return;
        }

        DidReset = true;
        Status = done.Count == 0
            ? Loc.Text("ResetDataWarning", "This action is permanent and cannot be undone.")
            : string.Join(", ", done);

        Completed?.Invoke();
    }
}
