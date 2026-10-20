using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Achievements;

namespace RustPlusDesk.Infrastructure.Achievements;

/// <summary>
/// Reference AchievementService's cache file: which achievements were earned and when, and which of them have not
/// been looked at yet. It sits beside the other local app data, so it travels with the install and not with the app.
/// </summary>
public sealed class AchievementFileStore(string path) : IAchievementStore
{
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "achievements.json");

    public AchievementState Load()
    {
        try
        {
            if (!File.Exists(path)) return new AchievementState();
            if (JsonNode.Parse(File.ReadAllBytes(path)) is not JsonObject document) return new AchievementState();

            var earned = new Dictionary<string, DateTime>();
            if (document["Earned"] is JsonObject entries)
            {
                foreach (var entry in entries)
                {
                    if (entry.Value is not JsonValue value) continue;
                    if (value.TryGetValue(out DateTime when)) earned[entry.Key] = when;
                    else if (value.TryGetValue(out string text)
                        && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out DateTime parsed))
                        earned[entry.Key] = parsed;
                }
            }

            var unseen = document["Unseen"] is JsonArray list
                ? list.Select(node => node?.GetValue<string>() ?? "").Where(id => id.Length > 0).ToArray()
                : Array.Empty<string>();
            return new AchievementState { Earned = earned, Unseen = unseen };
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            return new AchievementState();
        }
    }

    public void Save(AchievementState state)
    {
        var earned = new JsonObject();
        foreach (var entry in state.Earned)
            earned[entry.Key] = entry.Value.ToString("O", CultureInfo.InvariantCulture);
        var document = new JsonObject
        {
            ["Earned"] = earned,
            ["Unseen"] = new JsonArray(state.Unseen.Select(id => (JsonNode)id!).ToArray())
        };
        string? folder = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        File.WriteAllText(temporary, document.ToJsonString());
        File.Move(temporary, path, overwrite: true);
    }
}
