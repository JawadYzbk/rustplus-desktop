using System.Text.Json.Nodes;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile.CustomTimers and AlertCustomTimer in profiles.json. Timers are profile state, not
/// settings: chat creation, the ticking warnings and deletion all write them back, and the serialized, backed-up
/// profile writer adopts the latest document so a concurrent edit is not lost.
/// </summary>
public sealed class LegacyCustomTimerStore : ICustomTimerStore
{
    public async Task<CustomTimerSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        return bytes is null
            ? CustomTimerSettings.Defaults
            : Read(LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey));
    }

    public async Task SaveAsync(string profilesPath, string profileKey, CustomTimerSettings settings, CancellationToken ct) =>
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node["AlertCustomTimer"] = settings.AlertCustomTimer;
            node["CustomTimers"] = new JsonArray(settings.Timers.Select(timer => (JsonNode)new JsonObject
            {
                ["Id"] = timer.Id,
                ["Name"] = timer.Name,
                ["Command"] = timer.Command,
                ["EndTimeUtc"] = JsonValue.Create(timer.EndTimeUtc.ToUniversalTime()),
                ["OriginChannel"] = timer.OriginChannel?.ToString(),
                ["EnableCountdownAudio"] = timer.EnableCountdownAudio,
                ["EnableAlarmAudio"] = timer.EnableAlarmAudio,
                ["CreatedNotified"] = timer.CreatedNotified,
                ["Notified60"] = timer.Notified60,
                ["Notified30"] = timer.Notified30,
                ["Notified10"] = timer.Notified10,
                ["Notified3"] = timer.Notified3
            }).ToArray());
        }, null, ct).ConfigureAwait(false);

    private static CustomTimerSettings Read(JsonObject? profile)
    {
        var defaults = CustomTimerSettings.Defaults;
        if (profile is null) return defaults;
        bool alert = profile["AlertCustomTimer"] is JsonValue flag && flag.TryGetValue(out bool value) ? value : true;
        if (profile["CustomTimers"] is not JsonArray rows) return defaults with { AlertCustomTimer = alert };

        var timers = new List<ChatCustomTimer>();
        foreach (var row in rows)
        {
            if (row is not JsonObject timer) continue;
            string? id = Text(timer, "Id");
            string? name = Text(timer, "Name");
            string? command = Text(timer, "Command");
            DateTime? end = Timestamp(timer, "EndTimeUtc");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name)
                || string.IsNullOrWhiteSpace(command) || end is null) continue;
            timers.Add(new ChatCustomTimer(id, name, command, end.Value,
                OriginChannel(timer, "OriginChannel"),
                Flag(timer, "EnableCountdownAudio") ?? true,
                Flag(timer, "EnableAlarmAudio") ?? false,
                Flag(timer, "CreatedNotified") ?? false,
                Flag(timer, "Notified60") ?? false,
                Flag(timer, "Notified30") ?? false,
                Flag(timer, "Notified10") ?? false,
                Flag(timer, "Notified3") ?? false));
        }
        return new CustomTimerSettings(timers.Count == 0 ? CustomTimerSettings.NoTimers : timers, alert);
    }

    /// <summary>Reference OriginChannel is "Team", "Clan", or null for a timer made in the app.</summary>
    private static ChatChannel? OriginChannel(JsonObject timer, string name) => Text(timer, name) switch
    {
        "Team" => ChatChannel.Team,
        "Clan" => ChatChannel.Clan,
        _ => null
    };

    private static bool? Flag(JsonObject timer, string name) =>
        timer[name] is JsonValue value && value.TryGetValue(out bool flag) ? flag : null;

    private static string? Text(JsonObject timer, string name) =>
        timer[name] is JsonValue value && value.TryGetValue(out string? text) ? text : null;

    /// <summary>The reference stores a UTC DateTime; a value without a zone is read as UTC rather than local time.</summary>
    private static DateTime? Timestamp(JsonObject timer, string name)
    {
        if (timer[name] is not JsonValue value || !value.TryGetValue(out DateTime at)) return null;
        return at.Kind == DateTimeKind.Utc ? at : DateTime.SpecifyKind(at.ToUniversalTime(), DateTimeKind.Utc);
    }
}
