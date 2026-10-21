using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using RustPlusDesk.Core.Tutorials;

namespace RustPlusDesk.Avalonia.Features.Tutorials;

/// <summary>
/// Where the tutorial steps point.
///
/// The port has no attached property for this, so the shell says which control answers to which target id, and the
/// resolver asks here. A target that was never registered is a target that is not there, and the step is shown in the
/// middle with the honest "not available" wording rather than pointing at nothing.
/// </summary>
public sealed class TutorialTargetRegistry
{
    private readonly Dictionary<string, Func<Control?>> targets = new(StringComparer.Ordinal);

    /// <summary>Registers the control a target id points at.</summary>
    public void Register(string id, Func<Control?> resolver)
    {
        if (string.IsNullOrWhiteSpace(id)) return;
        targets[id] = resolver;
    }

    /// <summary>The control a target id points at, or null when there is none.</summary>
    public Control? Find(string id) => targets.TryGetValue(id, out var resolver) ? resolver() : null;

    /// <summary>The ids that have been registered.</summary>
    public IReadOnlyList<string> Registered => targets.Keys.ToList();
}

/// <summary>
/// Reference TutorialTargetResolver: finds the control a step points at and says where it is.
///
/// The bounds are in the window's own coordinates, which is the space the overlay draws in.
/// </summary>
public sealed class AvaloniaTutorialTargetResolver(TutorialTargetRegistry registry) : ITutorialTargetResolver
{
    public async Task<TutorialTarget?> ResolveAsync(TutorialStep step, CancellationToken cancellationToken = default)
    {
        if (step.TargetId is null) return null;

        var control = registry.Find(step.TargetId);
        if (control is null) return null;

        // The control has to have been laid out before it can be measured, and the step may have just opened the page
        // it points at. A turn of the dispatcher is what lets that layout happen.
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Loaded);

        if (control.Bounds.Width <= 0 || control.Bounds.Height <= 0) return null;

        var root = TopLevel.GetTopLevel(control);
        if (root is null) return null;

        var topLeft = control.TranslatePoint(new Point(0, 0), root);
        var bottomRight = control.TranslatePoint(new Point(control.Bounds.Width, control.Bounds.Height), root);
        if (topLeft is not { } start || bottomRight is not { } end) return null;

        double width = end.X - start.X;
        double height = end.Y - start.Y;
        if (width <= 0 || height <= 0) return null;

        return new TutorialTarget(new TutorialRect(start.X, start.Y, width, height), step.AllowTargetInteraction);
    }
}
