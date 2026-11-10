using CommunityToolkit.Mvvm.ComponentModel;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference Views/Windows/BackupPasswordDialog: the password a backup is encrypted with, or the one an encrypted
/// backup is opened with.
///
/// The reference derives the key from one field, and an empty one means "do not encrypt", so this window offers the
/// same single field and says what an empty one does - it does not ask for a confirmation the old app never asked for.
/// </summary>
public sealed partial class BackupPasswordViewModel(bool isDecrypting) : ObservableObject
{
    /// <summary>Whether the password opens a backup (true) or protects a new one (false).</summary>
    public bool IsDecrypting { get; } = isDecrypting;

    /// <summary>Reference TxtTitle.</summary>
    public string Title => Loc.Text(
        IsDecrypting ? "RestoreBackupTitle" : "BackupProtectionTitle",
        IsDecrypting ? "Restore Backup" : "Backup Protection");

    /// <summary>Reference TxtDescription.</summary>
    public string Description => Loc.Text(
        IsDecrypting ? "RestoreBackupDesc" : "BackupProtectionDesc",
        IsDecrypting
            ? "This backup is password-protected. Enter the password to decrypt it:"
            : "Enter a password to encrypt this backup file. If you leave it empty, the backup will not be encrypted.");

    /// <summary>Reference BtnConfirm's content.</summary>
    public string ConfirmLabel => Loc.Text(IsDecrypting ? "Decrypt" : "Encrypt", IsDecrypting ? "Decrypt" : "Encrypt");

    [ObservableProperty] private string password = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasError))]
    private string error = "";

    public bool HasError => Error.Length > 0;

    /// <summary>The password the backup is made with; an empty one means it is not encrypted.</summary>
    public string? EffectivePassword => Password.Length == 0 ? null : Password;

    /// <summary>
    /// Reference BtnConfirm_Click: an encrypted backup cannot be opened without a password, so the field must be
    /// filled in; a backup being written may leave it empty, which is how the old app makes an unencrypted one.
    /// </summary>
    public bool TryConfirm()
    {
        if (IsDecrypting && Password.Length == 0)
        {
            Error = Loc.Text("RestorePasswordErrorMessage", "Decryption failed. Invalid password or corrupted backup file.");
            return false;
        }

        Error = "";
        return true;
    }
}
