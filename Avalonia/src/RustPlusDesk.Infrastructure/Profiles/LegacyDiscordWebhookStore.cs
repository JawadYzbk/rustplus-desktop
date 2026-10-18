using System.Text.Json.Nodes;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile DiscordWebhookChatAlertsUrl/Mention/Enabled/Tts/Exclusive in profiles.json; writes go through
/// the serialized, backed-up profile writer, adopt the latest document and change only these five fields.
/// </summary>
public sealed class LegacyDiscordWebhookStore : IDiscordWebhookStore
{
    private const string Prefix = "DiscordWebhookChatAlerts";

    public async Task<DiscordWebhookSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        if (bytes is null || LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey) is not { } profile) return new();
        return new(Text(profile[Prefix + "Url"]), Text(profile[Prefix + "Mention"]), Flag(profile[Prefix + "Enabled"]), Flag(profile[Prefix + "Tts"]),
            Flag(profile[Prefix + "Exclusive"]));
    }

    public async Task SetAsync(string profilesPath, string profileKey, DiscordWebhookSettings settings, CancellationToken ct) =>
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node[Prefix + "Url"] = settings.Url;
            node[Prefix + "Mention"] = settings.Mention;
            node[Prefix + "Enabled"] = settings.Enabled;
            node[Prefix + "Tts"] = settings.Tts;
            node[Prefix + "Exclusive"] = settings.Exclusive;
        }, null, ct).ConfigureAwait(false);

    private static string Text(JsonNode? node) => node is JsonValue value && value.TryGetValue(out string? text) ? text : "";
    private static bool Flag(JsonNode? node) => node is JsonValue value && value.TryGetValue(out bool flag) && flag;
}
