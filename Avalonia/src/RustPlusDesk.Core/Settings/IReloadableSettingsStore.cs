namespace RustPlusDesk.Core.Settings;

/// <summary>
/// A settings store that can be read again while the app runs, for the flows that replace the settings file
/// underneath it - a restored backup above all.
///
/// It is separate from <see cref="IAppSettingsStore"/> so a store that has no file to read (the tests' in-memory one)
/// does not have to pretend otherwise.
/// </summary>
public interface IReloadableSettingsStore
{
    /// <summary>
    /// Adopts the values the file carries now. A change this app made and has not written yet is left alone: the
    /// restored file wins everywhere the app has nothing pending.
    /// </summary>
    void Reload();
}
