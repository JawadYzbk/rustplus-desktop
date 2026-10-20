using System.Text.Json;
using System.Text.Json.Serialization;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference AiCompanionStore: the AI companion's settings, kept in the app's data folder.
///
/// The reference protects the key with the Windows account's data protection, which does not exist on the other two
/// platforms the app now runs on. The port writes the file with the current user's permissions instead - the same
/// private file the app already uses for the companion token - so the key is still nobody else's on every platform.
/// </summary>
public sealed class AiSettingsStore(string path) : IAiSettingsStore
{
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "ai-companion.json");

    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull };

    public async Task<AiSettings> LoadAsync(CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
        if (bytes is null) return new AiSettings();
        try
        {
            return JsonSerializer.Deserialize<AiSettings>(bytes) ?? new AiSettings();
        }
        catch (JsonException)
        {
            // A file that cannot be read is a companion that is not set up, not a broken app.
            return new AiSettings();
        }
    }

    public async Task SaveAsync(AiSettings settings, CancellationToken ct)
    {
        byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
        byte[] data = JsonSerializer.SerializeToUtf8Bytes(settings, Options);
        await AtomicPrivateFile.WriteAsync(path, data, original, backup: false, ct).ConfigureAwait(false);
    }
}
