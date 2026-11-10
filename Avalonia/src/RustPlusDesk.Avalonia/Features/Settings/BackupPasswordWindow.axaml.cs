using Avalonia.Controls;
using global::Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference Views/Windows/BackupPasswordDialog: the password a backup is written with, or the one an encrypted
/// backup is opened with.
///
/// The window answers true when the player confirmed, and the caller then writes or restores the backup.
/// </summary>
public sealed partial class BackupPasswordWindow : Window
{
    public BackupPasswordWindow()
    {
        InitializeComponent();
        // Reference BackupPasswordDialog's PbPassword.Focus().
        Opened += (_, _) => this.FindControl<TextBox>("PasswordBox")?.Focus();
    }

    private void OnConfirm(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not BackupPasswordViewModel model) return;
        if (!model.TryConfirm()) return;
        Close(true);
    }

    private void OnCancel(object? sender, RoutedEventArgs e) => Close(false);
}
