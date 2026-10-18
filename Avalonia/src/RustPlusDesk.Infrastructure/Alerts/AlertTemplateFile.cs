using System.Text.Json;
using RustPlusDesk.Core.Alerts;

namespace RustPlusDesk.Infrastructure.Alerts;

/// <summary>
/// Reference custom_alerts.json in %APPDATA%/RustPlusDesk (culture → key → template), shared with the WPF app. Like the
/// reference, an unreadable file reads as no custom templates and a failed write is ignored.
/// </summary>
public sealed class AlertTemplateFile(string path) : IAlertTemplateStore
{
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "custom_alerts.json");
    private static readonly JsonSerializerOptions Indented = new() { WriteIndented = true };

    public Dictionary<string, Dictionary<string, string>> Load()
    {
        try
        {
            if (!File.Exists(path)) return [];
            return JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(File.ReadAllText(path)) ?? [];
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return []; }
    }

    public void Save(IReadOnlyDictionary<string, Dictionary<string, string>> overrides)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, JsonSerializer.Serialize(overrides, Indented));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }
}
