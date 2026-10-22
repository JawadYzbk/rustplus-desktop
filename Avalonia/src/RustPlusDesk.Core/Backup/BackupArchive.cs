using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;

namespace RustPlusDesk.Core.Backup;

/// <summary>
/// One thing a backup takes with it: the name it is stored under and where it lives right now.
///
/// The name is the reference's own file name wherever the two apps share the file, so a backup written by either app
/// carries the same entries and can be read by the other.
/// </summary>
public sealed record BackupItem(string EntryName, string SourcePath, bool IsFolder = false);

/// <summary>
/// Reference Services/Data/BackupDataModule: an archive of the app's data, behind a password when one is given.
///
/// The envelope is the reference's, byte for byte - the 13-byte signature, a 16-byte salt, a 16-byte IV and then
/// AES-256-CBC over the zip, with the key derived by PBKDF2 (SHA-256, 10000 iterations) - so a backup written by one
/// app opens in the other. The zip's entries are the reference's file names, and restoring maps them onto this app's
/// own paths.
/// </summary>
public static class BackupArchive
{
    /// <summary>Reference EncSig.</summary>
    public static byte[] Signature { get; } = Encoding.UTF8.GetBytes("RUST+DESK_ENC");

    /// <summary>Reference the 16-byte salt and IV.</summary>
    private const int SaltLength = 16;
    private const int IvLength = 16;

    /// <summary>Reference the 32-byte AES-256 key and its 10000 PBKDF2 rounds.</summary>
    private const int KeyLength = 32;
    private const int Iterations = 10000;

    /// <summary>Reference IsBackupEncrypted: whether the file starts with the encrypted backup's signature.</summary>
    public static bool IsEncrypted(string path)
    {
        try
        {
            if (!File.Exists(path)) return false;
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
            if (stream.Length < Signature.Length) return false;
            var header = new byte[Signature.Length];
            int read = stream.Read(header, 0, header.Length);
            return read == header.Length && header.AsSpan().SequenceEqual(Signature);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return false; }
    }

    /// <summary>
    /// Reference CreateBackup: the items are zipped under their entry names and the archive is written to
    /// <paramref name="zipPath"/>, encrypted when a password is given. Returns the entries that were found and stored.
    /// </summary>
    public static IReadOnlyList<string> Create(string zipPath, IReadOnlyList<BackupItem> items, string? password)
    {
        var stored = new List<string>();
        string plain = Path.Combine(Path.GetTempPath(), "rustplusdesk-backup-" + Guid.NewGuid().ToString("N") + ".zip");
        string? encrypted = null;

        try
        {
            using (var zip = ZipFile.Open(plain, ZipArchiveMode.Create))
            {
                foreach (var item in items)
                {
                    if (item.IsFolder)
                    {
                        if (!Directory.Exists(item.SourcePath)) continue;
                        string root = item.SourcePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                        foreach (string file in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories))
                        {
                            // A relative entry name, so the archive opens the same on every platform.
                            string relative = Path.GetRelativePath(root, file).Replace('\\', '/');
                            zip.CreateEntryFromFile(file, $"{item.EntryName}/{relative}", CompressionLevel.Optimal);
                        }
                        stored.Add(item.EntryName);
                    }
                    else if (File.Exists(item.SourcePath))
                    {
                        zip.CreateEntryFromFile(item.SourcePath, item.EntryName, CompressionLevel.Optimal);
                        stored.Add(item.EntryName);
                    }
                }
            }

            string source = plain;
            if (!string.IsNullOrEmpty(password))
            {
                encrypted = Path.Combine(Path.GetTempPath(), "rustplusdesk-backup-" + Guid.NewGuid().ToString("N") + ".enc");
                Encrypt(plain, encrypted, password);
                source = encrypted;
            }

            string? folder = Path.GetDirectoryName(Path.GetFullPath(zipPath));
            if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
            File.Copy(source, zipPath, overwrite: true);
            return stored;
        }
        finally
        {
            DeleteQuietly(plain);
            if (encrypted is not null) DeleteQuietly(encrypted);
        }
    }

    /// <summary>
    /// Reference RestoreBackup's unpacking: the archive is opened into <paramref name="destination"/>, decrypted first
    /// when it carries the signature. Returns the top-level entry names the archive holds.
    /// </summary>
    public static IReadOnlyList<string> Extract(string zipPath, string destination, string? password)
    {
        Directory.CreateDirectory(destination);
        string? decrypted = null;
        string source = zipPath;

        if (IsEncrypted(zipPath))
        {
            if (string.IsNullOrEmpty(password)) throw new InvalidDataException("The backup is encrypted and no password was given.");
            decrypted = Path.Combine(Path.GetTempPath(), "rustplusdesk-restore-" + Guid.NewGuid().ToString("N") + ".zip");
            Decrypt(zipPath, decrypted, password);
            source = decrypted;
        }

        try
        {
            using var zip = ZipFile.OpenRead(source);
            var entries = zip.Entries
                .Select(entry => entry.FullName.Replace('\\', '/'))
                .Where(name => name.Length > 0)
                .Select(name => name.Split('/', 2)[0])
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            zip.ExtractToDirectory(destination, overwriteFiles: true);
            return entries;
        }
        finally
        {
            if (decrypted is not null) DeleteQuietly(decrypted);
        }
    }

    /// <summary>Reference EncryptFile: the signature, the salt, the IV, then the encrypted bytes.</summary>
    private static void Encrypt(string sourcePath, string destinationPath, string password)
    {
        byte[] salt = RandomNumberGenerator.GetBytes(SaltLength);
        byte[] iv = RandomNumberGenerator.GetBytes(IvLength);
        byte[] key = Rfc2898DeriveBytes.Pbkdf2(password, salt, Iterations, HashAlgorithmName.SHA256, KeyLength);

        using var output = new FileStream(destinationPath, FileMode.Create, FileAccess.Write);
        output.Write(Signature);
        output.Write(salt);
        output.Write(iv);

        using var aes = Aes.Create();
        aes.Key = key;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        using var crypto = new CryptoStream(output, aes.CreateEncryptor(), CryptoStreamMode.Write);
        using var input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read);
        input.CopyTo(crypto);
    }

    /// <summary>Reference DecryptFile: the same envelope in reverse, and a wrong password fails here.</summary>
    private static void Decrypt(string sourcePath, string destinationPath, string password)
    {
        using var input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read);
        var header = new byte[Signature.Length];
        if (input.Read(header, 0, header.Length) != header.Length || !header.AsSpan().SequenceEqual(Signature))
            throw new InvalidDataException("The file is not a Rust+ Desktop backup.");

        var salt = new byte[SaltLength];
        var iv = new byte[IvLength];
        if (input.Read(salt, 0, salt.Length) != salt.Length || input.Read(iv, 0, iv.Length) != iv.Length)
            throw new InvalidDataException("The backup's envelope is incomplete.");

        byte[] key = Rfc2898DeriveBytes.Pbkdf2(password, salt, Iterations, HashAlgorithmName.SHA256, KeyLength);

        using var aes = Aes.Create();
        aes.Key = key;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        using var crypto = new CryptoStream(input, aes.CreateDecryptor(), CryptoStreamMode.Read);
        using var output = new FileStream(destinationPath, FileMode.Create, FileAccess.Write);
        crypto.CopyTo(output);
    }

    private static void DeleteQuietly(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* Left for the next sweep. */ }
    }
}
