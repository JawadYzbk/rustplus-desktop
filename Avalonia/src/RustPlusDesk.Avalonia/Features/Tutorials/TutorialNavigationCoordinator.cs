using RustPlusDesk.Core.Tutorials;

namespace RustPlusDesk.Avalonia.Features.Tutorials;

/// <summary>
/// Reference the tutorial's navigation: the page a step points at is opened before the step is shown, and whatever was
/// open before is put back when the tutorial ends.
/// </summary>
public sealed class TutorialNavigationCoordinator(Action<string> openPage) : ITutorialNavigationCoordinator
{
    /// <summary>Opens a page by its key, for example "settings" or "devices".</summary>
    public Action<string> OpenPage { get; set; } = openPage;

    /// <summary>What the page was before the tutorial opened another one, so it can be put back.</summary>
    private string? previousPage;

    /// <summary>The page that was open before the tutorial started.</summary>
    public string? CurrentPage { get; private set; }

    /// <summary>Opens the page a step points at.</summary>
    public Task PrepareAsync(TutorialStep step, CancellationToken cancellationToken = default)
    {
        if (step.PageKey is null) return Task.CompletedTask;

        previousPage ??= CurrentPage;
        CurrentPage = step.PageKey;
        OpenPage(step.PageKey);
        return Task.CompletedTask;
    }

    /// <summary>Remembers the page that was open, so the tutorial can put it back.</summary>
    public ITutorialStateSnapshot CaptureState() => new Snapshot(this, previousPage, CurrentPage);

    /// <summary>Puts the page that was open back.</summary>
    private void Restore(string? page)
    {
        if (page is null) return;
        CurrentPage = page;
        OpenPage(page);
    }

    /// <summary>
    /// The state a tutorial changes: the page that was open when it started.
    /// </summary>
    private sealed class Snapshot(TutorialNavigationCoordinator owner, string? previous, string? current) : ITutorialStateSnapshot
    {
        public Task RestoreAsync(CancellationToken cancellationToken = default)
        {
            owner.Restore(previous ?? current);
            return Task.CompletedTask;
        }
    }
}
