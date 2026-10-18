using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Console;

public sealed partial class ConsoleHelperView : UserControl
{
    public ConsoleHelperView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => { if (DataContext is ConsoleHelperViewModel model) model.CopyText = CopyAsync; };
    }

    private async Task CopyAsync(string text)
    {
        if (TopLevel.GetTopLevel(this)?.Clipboard is not { } clipboard) throw new InvalidOperationException("No clipboard.");
        await clipboard.SetTextAsync(text);
    }

    private void ItemPicked(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is not ListBox { SelectedItem: ConsoleItemChoice choice, DataContext: ConsoleParamViewModel param } list) return;
        param.ChooseItem(choice);
        list.SelectedItem = null;
    }

    /// <summary>Reference: a modal prompt reads one key, so no keystroke is read that the user did not offer.</summary>
    private async void CaptureKey(object? sender, RoutedEventArgs e)
    {
        if (sender is not Control { DataContext: ConsoleCommandViewModel command } || TopLevel.GetTopLevel(this) is not Window owner) return;
        var dialog = new KeyCaptureWindow(command.BindKey);
        if (await dialog.ShowDialog<bool>(owner)) command.BindKey = dialog.CapturedKey ?? "";
    }
}
