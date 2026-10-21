using System.Text.Json;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference AiCompanionStore: the AI companion's settings, kept in the app's data folder under the reference's file
/// name and JSON property names.
///
/// The reference protects each key with the Windows account's data protection, which does not exist on the other two
/// platforms the app now runs on. The port writes the file with the current user's permissions instead - the same
/// private file the app already uses for the companion token - so the key is still nobody else's on every platform.
///
/// Loads and saves are serialized: the settings card saves a model on every keystroke, and two writers racing on a
/// single small file would otherwise turn one of them into a lost update.
/// </summary>
public sealed class AiSettingsStore(string path) : IAiSettingsStore
{
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "ai-companion.json");

    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    private readonly SemaphoreSlim gate = new(1, 1);

    public async Task<AiSettings> LoadAsync(CancellationToken ct)
    {
        await gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            byte[]? bytes = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            if (bytes is null) return new AiSettings();

            AiSettings loaded;
            try
            {
                loaded = JsonSerializer.Deserialize<AiSettings>(bytes, Options) ?? new AiSettings();
            }
            catch (JsonException)
            {
                // A file that cannot be read is a companion that is not set up, not a broken app.
                return new AiSettings();
            }

            AiSettings migrated = loaded.MigrateLegacy();
            if (!ReferenceEquals(loaded, migrated))
            {
                // The old single Model/ApiKey slots have been moved into the per-provider dictionaries. Writing the
                // reference shape now means the migration runs once rather than on every start.
                byte[] data = JsonSerializer.SerializeToUtf8Bytes(migrated, Options);
                await AtomicPrivateFile.WriteAsync(path, data, bytes, backup: false, ct).ConfigureAwait(false);
            }
            return migrated;
        }
        finally
        {
            gate.Release();
        }
    }

    public async Task SaveAsync(AiSettings settings, CancellationToken ct)
    {
        await gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            byte[] data = JsonSerializer.SerializeToUtf8Bytes(settings, Options);
            await AtomicPrivateFile.WriteAsync(path, data, original, backup: false, ct).ConfigureAwait(false);
        }
        finally
        {
            gate.Release();
        }
    }
}
