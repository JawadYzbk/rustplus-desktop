using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Diagnostics;

namespace RustPlusDesk.Avalonia.Features.Activity;

/// <summary>
/// The bottom "Activity" log (reference LogPanel): filter box, Expand/Collapse between the reference 200 and
/// 420 px heights, and a read-only stream of timestamped, redacted lines.
/// </summary>
public sealed partial class ActivityLogViewModel(TimeProvider? clock = null) : ObservableObject
{
    public const double CollapsedHeight = 200, ExpandedHeight = 420;
    private readonly ActivityLog log = new(clock);

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Text))]
    [NotifyPropertyChangedFor(nameof(HasNoMatches))]
    private string filter = "";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ExpandCaption))]
    private bool isExpanded;

    public string Text => string.Join(Environment.NewLine, log.Filter(Filter));
    public bool IsEmpty => log.Lines.Count == 0;
    public bool HasNoMatches => !IsEmpty && !string.IsNullOrWhiteSpace(Filter) && !log.Filter(Filter).Any();
    public string ExpandCaption => IsExpanded ? "Collapse" : "Expand";
    public IReadOnlyList<string> Lines => log.Lines;

    public void Append(string? line)
    {
        if (!log.Append(line)) return;
        OnPropertyChanged(nameof(Text));
        OnPropertyChanged(nameof(IsEmpty));
        OnPropertyChanged(nameof(HasNoMatches));
    }

    [RelayCommand] private void ToggleExpand() => IsExpanded = !IsExpanded;
    [RelayCommand] private void ClearFilter() => Filter = "";
}
