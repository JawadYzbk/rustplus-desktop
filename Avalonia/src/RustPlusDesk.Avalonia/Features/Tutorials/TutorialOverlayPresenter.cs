using CommunityToolkit.Mvvm.ComponentModel;
using RustPlusDesk.Core.Tutorials;

namespace RustPlusDesk.Avalonia.Features.Tutorials;

/// <summary>
/// Reference the tutorial overlay's presenter: the step that is on screen, and the fact that something is.
///
/// The core builds the presentation and the overlay only draws it, so the state machine never has to know what a
/// window looks like.
/// </summary>
public sealed partial class TutorialOverlayPresenter : ObservableObject, ITutorialPresenter
{
    /// <summary>The step that is being shown, or null when no tutorial is running.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsVisible))]
    [NotifyPropertyChangedFor(nameof(Title))]
    [NotifyPropertyChangedFor(nameof(StepTitle))]
    [NotifyPropertyChangedFor(nameof(Description))]
    [NotifyPropertyChangedFor(nameof(Tip))]
    [NotifyPropertyChangedFor(nameof(StepText))]
    [NotifyPropertyChangedFor(nameof(Placement))]
    [NotifyPropertyChangedFor(nameof(HasTarget))]
    [NotifyPropertyChangedFor(nameof(AllowInteraction))]
    [NotifyPropertyChangedFor(nameof(CanGoBack))]
    [NotifyPropertyChangedFor(nameof(IsLastStep))]
    [NotifyPropertyChangedFor(nameof(IsTargetUnavailable))]
    private TutorialPresentation? presentation;

    /// <summary>Whether the overlay is showing a step.</summary>
    public bool IsVisible => Presentation is not null;

    /// <summary>The tutorial's name.</summary>
    public string Title => Presentation?.TutorialTitle ?? "";

    /// <summary>The step's name.</summary>
    public string StepTitle => Presentation?.StepTitle ?? "";

    /// <summary>What the step says.</summary>
    public string Description => Presentation?.Description ?? "";

    /// <summary>The tip under the description, if the step has one.</summary>
    public string? Tip => Presentation?.Tip;

    /// <summary>How far through the tutorial the player is.</summary>
    public string StepText => Presentation is null
        ? ""
        : string.Format(RustPlusDesk.Core.Localization.Loc.Text("Tutorials.Common.StepCount", "Step {0} of {1}"),
            Presentation.StepNumber, Presentation.StepCount);

    /// <summary>Where the card goes relative to the target.</summary>
    public TutorialPlacement Placement => Presentation?.Placement ?? TutorialPlacement.Center;

    /// <summary>Whether there is a target to point at.</summary>
    public bool HasTarget => Presentation is { Target: { } target } && !target.Bounds.IsEmpty;

    /// <summary>Whether the target can be used while the spotlight is on it.</summary>
    public bool AllowInteraction => Presentation?.AllowTargetInteraction == true;

    /// <summary>Whether the step before this one can be shown again.</summary>
    public bool CanGoBack => Presentation?.CanGoBack == true;

    /// <summary>Whether this is the last step of the tutorial.</summary>
    public bool IsLastStep => Presentation?.IsLastStep == true;

    /// <summary>Whether the target could not be found, so the step is shown in the middle instead.</summary>
    public bool IsTargetUnavailable => Presentation?.IsTargetUnavailable == true;

    /// <summary>The step that is on screen, for the overlay to draw.</summary>
    public void Show(TutorialPresentation presentation)
    {
        Presentation = presentation;
        OnPropertyChanged(nameof(IsVisible));
    }

    /// <summary>The tutorial is over, or was cancelled.</summary>
    public void Hide()
    {
        Presentation = null;
        OnPropertyChanged(nameof(IsVisible));
    }
}
