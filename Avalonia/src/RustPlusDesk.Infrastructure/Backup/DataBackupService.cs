using RustPlusDesk.Core.Backup;

namespace RustPlusDesk.Infrastructure.Backup;

/// <summary>
/// Reference BackupDataModule: writes a backup of the app's data and puts one back.
///
/// The archive is built and unpacked through <see cref="BackupArchive"/>, and this class is the part that knows where
/// this app keeps its files - so restoring onto another machine, or from a backup the old app wrote, puts the files
/// where this app reads them rather than where they were.
/// </summary>
public sealed class DataBackupService(
    Func<string?> profilesPath,
    string? appFolder = null,
    string? cacheFolder = null,
    string? localFolder = null,
    string? shopAlertsPath = null)
{
    /// <summary>Everything a backup of this machine would carry, and where each entry belongs on it.</summary>
    public IReadOnlyList<BackupEntry> Plan() => AppDataBackupPlan.Entries(
        string.IsNullOrWhiteSpace(profilesPath()) ? AppDataBackupPlan.DefaultProfilesPath : profilesPath()!,
        appFolder, cacheFolder, localFolder, shopAlertsPath);

    /// <summary>
    /// Reference CreateBackup: the entries that exist are zipped to <paramref name="zipPath"/>, encrypted when a
    /// password is given. Returns the entries the backup carries.
    /// </summary>
    public IReadOnlyList<string> Create(string zipPath, string? password) =>
        BackupArchive.Create(zipPath, [.. Plan().Select(entry => new BackupItem(entry.EntryName, entry.Path, entry.IsFolder))], password);

    /// <summary>
    /// Reference RestoreBackup: the archive is unpacked aside and each entry it carries is put back where this app
    /// keeps it, replacing what is there now. Returns the entries that were restored.
    /// </summary>
    public IReadOnlyList<string> Restore(string zipPath, string? password)
    {
        string staging = Path.Combine(Path.GetTempPath(), "rustplusdesk-restore-" + Guid.NewGuid().ToString("N"));
        try
        {
            BackupArchive.Extract(zipPath, staging, password);
            var restored = new List<string>();

            foreach (var entry in Plan())
            {
                string source = Path.Combine(staging, entry.EntryName.Replace('/', Path.DirectorySeparatorChar));
                if (entry.IsFolder)
                {
                    if (!Directory.Exists(source)) continue;
                    CopyFolder(source, entry.Path);
                    restored.Add(entry.EntryName);
                }
                else
                {
                    if (!File.Exists(source)) continue;
                    string? folder = Path.GetDirectoryName(Path.GetFullPath(entry.Path));
                    if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
                    File.Copy(source, entry.Path, overwrite: true);
                    restored.Add(entry.EntryName);
                }
            }

            return restored;
        }
        finally
        {
            try { if (Directory.Exists(staging)) Directory.Delete(staging, recursive: true); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* Left for the next sweep. */ }
        }
    }

    private static void CopyFolder(string source, string destination)
    {
        Directory.CreateDirectory(destination);
        foreach (string file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
        {
            string target = Path.Combine(destination, Path.GetRelativePath(source, file));
            string? folder = Path.GetDirectoryName(target);
            if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
            File.Copy(file, target, overwrite: true);
        }
    }
}
