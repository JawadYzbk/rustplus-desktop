using RustPlusDesk.Core.Tutorials;

namespace RustPlusDesk.Avalonia.Features.Tutorials;

/// <summary>
/// Reference the app state a tutorial step's condition reads.
///
/// The port answers the same questions from the view models it has. Where it has no equivalent the answer is the
/// honest one - a feature that is not there is not offered, and the step that needs it is skipped rather than shown
/// pointing at nothing.
/// </summary>
public sealed class TutorialContext : ITutorialContext
{
    /// <summary>Whether the player is signed in to the cloud account.</summary>
    public Func<bool> IsLoggedIn { get; set; } = () => false;

    /// <summary>Whether the account is a supporter account.</summary>
    public Func<bool> IsPremium { get; set; } = () => false;

    /// <summary>Whether a server is paired.</summary>
    public Func<bool> HasPairedServer { get; set; } = () => false;

    /// <summary>Whether a server is selected.</summary>
    public Func<bool> HasSelectedServer { get; set; } = () => false;

    /// <summary>Whether the app is softly connected to the server.</summary>
    public Func<bool> IsSoftConnected { get; set; } = () => false;

    /// <summary>Whether the app is fully connected to the server.</summary>
    public Func<bool> IsFullConnected { get; set; } = () => false;

    /// <summary>Whether the map is loaded.</summary>
    public Func<bool> HasMap { get; set; } = () => false;

    /// <summary>Whether there are paired devices.</summary>
    public Func<bool> HasDevices { get; set; } = () => false;

    /// <summary>Whether there is a team.</summary>
    public Func<bool> HasTeam { get; set; } = () => false;

    /// <summary>Whether a Discord webhook is set up.</summary>
    public Func<bool> IsDiscordConfigured { get; set; } = () => false;

    /// <summary>Whether there are automation rules.</summary>
    public Func<bool> HasAutomationRules { get; set; } = () => false;

    bool ITutorialContext.IsLoggedIn => IsLoggedIn();
    bool ITutorialContext.IsPremium => IsPremium();
    bool ITutorialContext.HasPairedServer => HasPairedServer();
    bool ITutorialContext.HasSelectedServer => HasSelectedServer();
    bool ITutorialContext.IsSoftConnected => IsSoftConnected();
    bool ITutorialContext.IsFullConnected => IsFullConnected();
    bool ITutorialContext.HasMap => HasMap();
    bool ITutorialContext.HasDevices => HasDevices();
    bool ITutorialContext.HasTeam => HasTeam();
    bool ITutorialContext.IsDiscordConfigured => IsDiscordConfigured();
    bool ITutorialContext.HasAutomationRules => HasAutomationRules();
}
