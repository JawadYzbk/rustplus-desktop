using System.Text.Json.Serialization;

namespace RustPlusDesk.Core.Tutorials;

/// <summary>Reference TutorialPlacement: where the instruction card is put relative to its target.</summary>
public enum TutorialPlacement { Auto, Top, Bottom, Left, Right, Center }

/// <summary>Reference TutorialStatus: what the saved progress of one tutorial says.</summary>
public enum TutorialStatus { NotStarted, InProgress, Completed, Skipped, Updated }

/// <summary>
/// Reference System.Windows.Thickness, kept UI-free so the tutorial catalogue can live in the core: the distance
/// the spotlight is grown by on each side of its target.
/// </summary>
public readonly record struct TutorialPadding(double Left, double Top, double Right, double Bottom)
{
    public TutorialPadding(double uniform) : this(uniform, uniform, uniform, uniform) { }

    public static TutorialPadding Empty => default;

    public double Horizontal => Left + Right;
    public double Vertical => Top + Bottom;
}

/// <summary>
/// Reference System.Windows.Rect, kept UI-free: a target's bounds in the overlay's coordinate space.
/// </summary>
public readonly record struct TutorialRect(double X, double Y, double Width, double Height)
{
    public static TutorialRect Empty => default;

    public bool IsEmpty => Width <= 0 || Height <= 0;
    public double Right => X + Width;
    public double Bottom => Y + Height;

    /// <summary>Grows the rectangle by the spotlight padding (reference TutorialTargetResolver.ApplyPadding).</summary>
    public TutorialRect Inflate(TutorialPadding padding) => new(
        X - padding.Left, Y - padding.Top, Width + padding.Horizontal, Height + padding.Vertical);
}

/// <summary>Reference TutorialDefinition: one chapter of the tutorial centre.</summary>
public sealed class TutorialDefinition
{
    public required string Id { get; init; }
    public required int Version { get; init; }
    public required string Category { get; init; }
    public required string TitleKey { get; init; }
    public required string DescriptionKey { get; init; }
    public string Icon { get; init; } = "Book24";
    public int DisplayOrder { get; init; }
    public bool IsRecommended { get; init; }
    public bool IsNewFeature { get; init; }
    public IReadOnlyList<TutorialStep> Steps { get; init; } = [];
}

/// <summary>Reference TutorialStep: one instruction, its target and the conditions that make it applicable.</summary>
public sealed class TutorialStep
{
    public required string Id { get; init; }
    public required string TitleKey { get; init; }
    public required string DescriptionKey { get; init; }
    public string? TipKey { get; init; }

    /// <summary>
    /// Reference TutorialStep.ImagePath: an illustration shown under the description for steps that explain
    /// something outside the app. The Avalonia port has no screenshot pack, so the overlay omits a missing image.
    /// </summary>
    public string? ImagePath { get; init; }
    public string? PageKey { get; init; }
    public string? TargetId { get; init; }

    /// <summary>Reference TutorialStep.WebViewTargetId: a target inside the 3D map's web view.</summary>
    public string? WebViewTargetId { get; init; }
    public TutorialPlacement Placement { get; init; } = TutorialPlacement.Auto;
    public TutorialPadding SpotlightPadding { get; init; } = new(8);
    public bool IsOptional { get; init; }
    public bool AllowTargetInteraction { get; init; }
    public bool AutoScrollIntoView { get; init; } = true;
    public Func<ITutorialContext, bool>? CanShow { get; init; }
    public Func<ITutorialContext, CancellationToken, Task>? BeforeShowAsync { get; init; }
    public Func<ITutorialContext, CancellationToken, Task>? AfterHideAsync { get; init; }
}

/// <summary>
/// Reference TutorialProgress: the saved state of one tutorial. Unknown members survive a round-trip through the
/// store, so a newer build (or the reference app) can add fields without this one dropping them.
/// </summary>
public sealed class TutorialProgress
{
    public string TutorialId { get; set; } = "";
    public int TutorialVersion { get; set; }
    public TutorialStatus Status { get; set; }
    public string? LastCompletedStepId { get; set; }
    public HashSet<string> CompletedStepIds { get; set; } = [];
    public DateTime? StartedAtUtc { get; set; }
    public DateTime? CompletedAtUtc { get; set; }
    public DateTime? SkippedAtUtc { get; set; }

    [JsonExtensionData]
    public Dictionary<string, System.Text.Json.JsonElement>? AdditionalData { get; set; }
}

/// <summary>Reference TutorialPreferences: the switches that decide whether the app offers its tours.</summary>
public sealed class TutorialPreferences
{
    public bool FirstRunPromptDismissed { get; set; }
    public bool AutoStartBasicTutorial { get; set; } = true;
    public bool AutoStartNewFeatureTutorials { get; set; } = true;
    public HashSet<string> OfferedTutorialIds { get; set; } = [];
    public string? LastTutorialId { get; set; }

    [JsonExtensionData]
    public Dictionary<string, System.Text.Json.JsonElement>? AdditionalData { get; set; }
}

/// <summary>
/// Reference TutorialTarget: the area the spotlight covers. A target with empty bounds means "show the step in the
/// centre" - a target that is honestly unavailable in this port, rather than a missing one.
/// </summary>
public sealed record TutorialTarget(TutorialRect Bounds, bool AllowInteraction = false)
{
    public static TutorialTarget Empty { get; } = new(TutorialRect.Empty);
}

/// <summary>Reference TutorialEventArgs: the tutorial, step and reason an event is about.</summary>
public sealed class TutorialEventArgs(string tutorialId, string? stepId = null, string? reason = null) : EventArgs
{
    public string TutorialId { get; } = tutorialId;
    public string? StepId { get; } = stepId;
    public string? Reason { get; } = reason;
}

/// <summary>
/// Reference ITutorialContext: the app state a step's condition reads. This port answers the same questions from
/// the view models it has; where it has no equivalent (a separate full connection, for example) the answer is the
/// nearest honest one and the gap is reported.
/// </summary>
public interface ITutorialContext
{
    bool IsLoggedIn { get; }
    bool IsPremium { get; }
    bool HasPairedServer { get; }
    bool HasSelectedServer { get; }
    bool IsSoftConnected { get; }
    bool IsFullConnected { get; }
    bool HasMap { get; }
    bool HasDevices { get; }
    bool HasTeam { get; }
    bool IsDiscordConfigured { get; }
    bool HasAutomationRules { get; }
}

/// <summary>Reference ITutorialStateSnapshot: the workspace state a tutorial changes, restored when it ends.</summary>
public interface ITutorialStateSnapshot
{
    Task RestoreAsync(CancellationToken cancellationToken = default);
}

/// <summary>The snapshot used when a tutorial has not changed anything yet.</summary>
public sealed class EmptyTutorialStateSnapshot : ITutorialStateSnapshot
{
    public static EmptyTutorialStateSnapshot Instance { get; } = new();
    public Task RestoreAsync(CancellationToken cancellationToken = default) => Task.CompletedTask;
}

/// <summary>Reference ITutorialNavigationCoordinator: opens the page a step points at and restores what was open.</summary>
public interface ITutorialNavigationCoordinator
{
    Task PrepareAsync(TutorialStep step, CancellationToken cancellationToken = default);
    ITutorialStateSnapshot CaptureState();
}

/// <summary>
/// Reference TutorialPresentation: everything the overlay needs to draw one step. The core builds it; the
/// Avalonia overlay only renders it.
/// </summary>
public sealed record TutorialPresentation(
    string TutorialTitle, string StepTitle, string Description, string? Tip,
    int StepNumber, int StepCount, TutorialPlacement Placement, TutorialTarget Target,
    bool CanGoBack, bool IsLastStep, bool IsTargetUnavailable, bool AllowTargetInteraction,
    string? ImagePath = null);

/// <summary>Reference ITutorialPresenter: the control that draws the tutorial.</summary>
public interface ITutorialPresenter
{
    bool IsVisible { get; }
    void Show(TutorialPresentation presentation);
    void Hide();
}

/// <summary>Reference ITutorialTargetResolver: finds a step's target and reports its bounds.</summary>
public interface ITutorialTargetResolver
{
    Task<TutorialTarget?> ResolveAsync(TutorialStep step, CancellationToken cancellationToken = default);
}

/// <summary>Reference ITutorialService: starts, advances, ends and resets the tutorial run.</summary>
public interface ITutorialService
{
    bool IsRunning { get; }
    TutorialDefinition? ActiveTutorial { get; }
    TutorialStep? ActiveStep { get; }
    event EventHandler<TutorialEventArgs>? TutorialStarted;
    event EventHandler<TutorialEventArgs>? StepShown;
    event EventHandler<TutorialEventArgs>? StepCompleted;
    event EventHandler<TutorialEventArgs>? StepSkipped;
    event EventHandler<TutorialEventArgs>? TutorialCompleted;
    event EventHandler<TutorialEventArgs>? TutorialSkipped;
    event EventHandler<TutorialEventArgs>? TutorialCancelled;
    event EventHandler<TutorialEventArgs>? TargetResolutionFailed;
    Task StartAsync(string tutorialId, string? startStepId = null, CancellationToken cancellationToken = default);
    Task ContinueAsync(string tutorialId, CancellationToken cancellationToken = default);
    Task NextAsync();
    Task BackAsync();
    Task SkipAsync();
    Task FinishAsync();
    Task CancelAsync();
    Task StartForCurrentPageAsync();
    Task ResetAsync(string tutorialId);
    Task ResetAllAsync();
    Task RefreshAsync();
}
