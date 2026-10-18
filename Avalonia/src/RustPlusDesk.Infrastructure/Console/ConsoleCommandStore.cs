using System.Text.Json;
using RustPlusDesk.Core.Console;

namespace RustPlusDesk.Infrastructure.Console;

/// <summary>
/// The bundled reference console-commands.json and the reference console-helper.json user state (same path and
/// shape: command ID → bind and changed values). An unreadable state file loads as empty, like the reference, but is
/// copied to a backup before the first save replaces it.
/// </summary>
public sealed class ConsoleCommandStore(string statePath) : IConsoleCommandStore
{
    private static readonly JsonSerializerOptions CatalogOptions = new()
    {
        PropertyNameCaseInsensitive = true, ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true
    };
    private static readonly JsonSerializerOptions StateOptions = new() { WriteIndented = true };
    private readonly SemaphoreSlim writes = new(1, 1);
    private bool unreadable;

    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RustPlusDesk", "console-helper.json");

    public ConsoleCommandCatalog LoadCatalog()
    {
        using var stream = typeof(ConsoleCommandStore).Assembly.GetManifestResourceStream("RustPlusDesk.Infrastructure.ConsoleCommands.json")
            ?? throw new InvalidDataException("The bundled console command list is missing.");
        return ParseCatalog(stream);
    }

    public static ConsoleCommandCatalog ParseCatalog(Stream stream) =>
        JsonSerializer.Deserialize<ConsoleCommandCatalog>(stream, CatalogOptions) ?? throw new InvalidDataException("The console command list is empty.");

    public async Task<IReadOnlyDictionary<string, ConsoleCommandUserState>> LoadStateAsync(CancellationToken ct)
    {
        if (!File.Exists(statePath)) return new Dictionary<string, ConsoleCommandUserState>();
        try
        {
            await using var stream = File.OpenRead(statePath);
            return await JsonSerializer.DeserializeAsync<Dictionary<string, ConsoleCommandUserState>>(stream, StateOptions, ct).ConfigureAwait(false)
                ?? new Dictionary<string, ConsoleCommandUserState>();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            unreadable = true;
            return new Dictionary<string, ConsoleCommandUserState>();
        }
    }

    public async Task SaveStateAsync(IReadOnlyDictionary<string, ConsoleCommandUserState> state, CancellationToken ct)
    {
        await writes.WaitAsync(ct).ConfigureAwait(false);
        string temporary = statePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(statePath))!);
            if (unreadable && File.Exists(statePath))
            {
                File.Copy(statePath, statePath + ".unreadable-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + ".bak", overwrite: true);
                unreadable = false;
            }
            await using (var stream = File.Create(temporary))
                await JsonSerializer.SerializeAsync(stream, state, StateOptions, ct).ConfigureAwait(false);
            File.Move(temporary, statePath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
            writes.Release();
        }
    }
}
