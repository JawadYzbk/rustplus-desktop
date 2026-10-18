using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Shared.Icons;

namespace RustPlusDesk.Avalonia.Features.Toasts;

/// <summary>Reference ControlAppearance severities used for the toast accent.</summary>
public enum ToastKind { Info, Success, Caution, Danger }

/// <summary>One toast (reference ToastItem): title, message, severity, icon and how long it stays.</summary>
public sealed partial class ToastItem : ObservableObject
{
    public string Title { get; init; } = "";
    public string Message { get; init; } = "";
    public ToastKind Kind { get; init; } = ToastKind.Info;
    public FluentSymbol Symbol { get; init; } = FluentSymbol.Info24;
    /// <summary>Zero pins the toast until it is dismissed.</summary>
    public TimeSpan Timeout { get; init; } = ToastsViewModel.DefaultTimeout;
    /// <summary>Runs once when the toast leaves for any reason.</summary>
    public Action? Closed { get; init; }
    public bool HasTitle => Title.Length > 0;
    public bool IsSuccess => Kind == ToastKind.Success;
    public bool IsCaution => Kind == ToastKind.Caution;
    public bool IsDanger => Kind == ToastKind.Danger;
    internal IDisposable? Timer { get; set; }
}

/// <summary>
/// The reference stacked toasts (MainWindow AddToast/DismissToast/ShowInfoSnackbar): newest at the bottom, at most five
/// (the oldest is evicted), eight seconds by default, a close button, and a Closed callback that always runs once.
/// </summary>
public sealed partial class ToastsViewModel(Func<TimeSpan, Action, IDisposable>? schedule = null) : ObservableObject
{
    public const int MaxVisible = 5;
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(8);
    private readonly Func<TimeSpan, Action, IDisposable> schedule = schedule ?? ((delay, action) => DispatcherTimer.RunOnce(action, delay));

    public ObservableCollection<ToastItem> Items { get; } = [];

    /// <summary>Reference ShowInfoSnackbar.</summary>
    public ToastItem Show(string title, string message, ToastKind kind = ToastKind.Info, FluentSymbol? symbol = null)
    {
        var item = new ToastItem { Title = title, Message = message, Kind = kind, Symbol = symbol ?? SymbolFor(kind) };
        Add(item);
        return item;
    }

    public void Add(ToastItem item)
    {
        while (Items.Count >= MaxVisible) Remove(Items[0]);
        Items.Add(item);
        if (item.Timeout > TimeSpan.Zero) item.Timer = schedule(item.Timeout, () => Dismiss(item));
    }

    [RelayCommand]
    public void Dismiss(ToastItem? item)
    {
        if (item is not null && Items.Contains(item)) Remove(item);
    }

    private void Remove(ToastItem item)
    {
        item.Timer?.Dispose();
        item.Timer = null;
        if (!Items.Remove(item)) return;
        item.Closed?.Invoke();
    }

    public static FluentSymbol SymbolFor(ToastKind kind) => kind switch
    {
        ToastKind.Success => FluentSymbol.CheckmarkCircle24,
        ToastKind.Caution => FluentSymbol.Warning24,
        ToastKind.Danger => FluentSymbol.AlertUrgent24,
        _ => FluentSymbol.Info24
    };
}
