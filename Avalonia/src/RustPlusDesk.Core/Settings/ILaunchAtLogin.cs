namespace RustPlusDesk.Core.Settings;

/// <summary>
/// Starts the app when the user signs in to the operating system (reference TrackingService.SetAutoStart, which
/// wrote the HKCU Run value "RustPlusDesk" = "\"exe\" --background"). Each OS gets its own registration.
/// </summary>
public interface ILaunchAtLogin
{
    /// <summary>False where this OS or build cannot register a login item; <see cref="UnsupportedReason"/> says why.</summary>
    bool IsSupported { get; }
    string UnsupportedReason { get; }
    /// <summary>Short OS-specific label, e.g. "Launch with Windows".</summary>
    string Label { get; }
    /// <summary>Adds or removes the login registration; returns false (and leaves things as they were) on failure.</summary>
    bool Apply(bool enabled);
}
