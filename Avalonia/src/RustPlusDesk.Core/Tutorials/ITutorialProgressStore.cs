namespace RustPlusDesk.Core.Tutorials;

/// <summary>
/// Reference ITutorialProgressStore: the saved progress and preferences of the tutorial centre. The reference keeps
/// them in <c>%APPDATA%\RustPlusDesk\tutorial-progress.json</c>; the implementation lives in the infrastructure.
/// </summary>
public interface ITutorialProgressStore
{
    Task<TutorialProgress> GetAsync(TutorialDefinition definition, CancellationToken cancellationToken = default);
    Task<IReadOnlyDictionary<string, TutorialProgress>> GetAllAsync(CancellationToken cancellationToken = default);
    Task SaveAsync(TutorialProgress progress, CancellationToken cancellationToken = default);
    Task ResetAsync(string tutorialId, CancellationToken cancellationToken = default);
    Task ResetAllAsync(CancellationToken cancellationToken = default);
    Task<TutorialPreferences> GetPreferencesAsync(CancellationToken cancellationToken = default);
    Task SavePreferencesAsync(TutorialPreferences preferences, CancellationToken cancellationToken = default);
}
