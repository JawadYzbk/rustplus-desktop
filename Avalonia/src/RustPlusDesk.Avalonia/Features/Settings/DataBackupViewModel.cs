using System.Globalization;
using System.Security.Cryptography;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Backup;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Infrastructure.Backup;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference BtnBackupData_Click and BtnRestoreData_Click: writing a backup of the app's data, and putting one back.
///
/// The service moves the files; this holds what the settings card shows - the confirmation the reference asks before
/// it even picks a file, and the line that says how the last backup or restore went. The file pickers belong to the
/// view, because only a window can open one.
/// </summary>
public sealed partial class DataBackupViewModel(
    DataBackupService service,
    Action<string>? log = null,
    Func<Task>? reloadAfterRestore = null) : ObservableObject
{
    /// <summary>How the last backup or restore went, in the reference's own words.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasStatus))]
    private string status = "";

    /// <summary>Whether a backup or a restore is running, so the card's buttons can wait.</summary>
    [ObservableProperty] private bool isWorking;

    /// <summary>Whether the restore confirmation is on the screen (reference MessageBoxButton.YesNo).</summary>
    [ObservableProperty] private bool hasRestoreConfirmation;

    public bool HasStatus => Status.Length > 0;

    /// <summary>Reference RestoreConfirmMessage: a restore asks before it overwrites anything.</summary>
    public string RestoreConfirmationText => Loc.Text("RestoreConfirmMessage",
        "Restoring data will overwrite your current server profiles, custom crosshairs, drawing overlays, and minimap configurations.\n\nAre you sure you want to continue?");

    /// <summary>Reference the save dialog's FileName: the backup's name carries the moment it was made.</summary>
    public string SuggestedFileName =>
        "RustPlusDesk_Backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".zip";

    /// <summary>Reference BackupDataModule.IsBackupEncrypted: whether the picked file asks for a password.</summary>
    public static bool IsEncrypted(string path) => BackupArchive.IsEncrypted(path);

    /// <summary>The restore button opens the confirmation first, as the reference's message box does.</summary>
    [RelayCommand]
    private void RequestRestore()
    {
        Status = "";
        HasRestoreConfirmation = true;
    }

    [RelayCommand] private void CancelRestore() => HasRestoreConfirmation = false;

    /// <summary>Raised once the player confirmed; only the view can open the picker.</summary>
    public event Action? RestoreConfirmed;

    [RelayCommand]
    private void ConfirmRestore()
    {
        HasRestoreConfirmation = false;
        RestoreConfirmed?.Invoke();
    }

    /// <summary>
    /// Reference CreateBackup: the archive is written to the path the player picked, encrypted when a password was
    /// given. The work runs off the UI thread and the card reports what happened afterwards.
    /// </summary>
    public async Task BackupAsync(string zipPath, string? password)
    {
        IsWorking = true;
        Status = "";
        try
        {
            await Task.Run(() => service.Create(zipPath, password)).ConfigureAwait(true);
            log?.Invoke(string.Format(CultureInfo.CurrentCulture,
                Loc.Text("BackupSuccessLog", "[SYSTEM] Backup created successfully at: {0}"), zipPath));
            Status = Loc.Text("BackupSuccessMessage", "Backup created successfully!");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException or CryptographicException)
        {
            log?.Invoke(string.Format(CultureInfo.CurrentCulture,
                Loc.Text("BackupErrorLog", "[SYSTEM] Backup error: {0}"), ex.Message));
            Status = string.Format(CultureInfo.CurrentCulture,
                Loc.Text("BackupErrorMessage", "Error creating backup: {0}"), ex.Message);
        }
        finally
        {
            IsWorking = false;
        }
    }

    /// <summary>
    /// Reference RestoreBackup and ReloadApplicationData: the archive is unpacked over this machine's data and the
    /// app reads the restored files again. A file that will not decrypt is reported as the reference reports it - the
    /// password and a corrupt archive share one message, because the two cannot be told apart.
    /// </summary>
    public async Task RestoreAsync(string zipPath, string? password)
    {
        IsWorking = true;
        Status = "";
        try
        {
            await Task.Run(() => service.Restore(zipPath, password)).ConfigureAwait(true);
            if (reloadAfterRestore is not null) await reloadAfterRestore().ConfigureAwait(true);
            Status = Loc.Text("RestoreSuccessMessage", "Application data restored successfully!");
        }
        catch (Exception ex) when (ex is CryptographicException or InvalidDataException)
        {
            log?.Invoke(Loc.Text("RestorePasswordErrorLog", "[SYSTEM] Restore error: Invalid password provided."));
            Status = Loc.Text("RestorePasswordErrorMessage", "Decryption failed. Invalid password or corrupted backup file.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            log?.Invoke(string.Format(CultureInfo.CurrentCulture,
                Loc.Text("RestoreErrorLog", "[SYSTEM] Restore error: {0}"), ex.Message));
            Status = string.Format(CultureInfo.CurrentCulture,
                Loc.Text("RestoreErrorMessage", "Error restoring backup: {0}"), ex.Message);
        }
        finally
        {
            IsWorking = false;
        }
    }
}
