namespace RustPlusDesk.Core.Alerts;

/// <summary>
/// A server profile's Discord webhook for chat alerts (reference ServerProfile.DiscordWebhookChatAlerts{Url, Mention,
/// Enabled, Tts, Exclusive}). Exclusive ("Discord only") keeps alerts out of in-game chat.
/// </summary>
public sealed record DiscordWebhookSettings(string Url = "", string Mention = "", bool Enabled = false, bool Tts = false, bool Exclusive = false)
{
    public const string Everyone = "@everyone", Here = "@here";

    /// <summary>Reference URL setter: an empty URL turns the webhook off, any other URL turns it on.</summary>
    public DiscordWebhookSettings WithUrl(string? url) => this with { Url = url ?? "", Enabled = !string.IsNullOrWhiteSpace(url) };

    public bool MentionsEveryone => Mention.Contains(Everyone, StringComparison.Ordinal);
    public bool MentionsHere => Mention.Contains(Here, StringComparison.Ordinal);

    /// <summary>Reference DiscordWebhookChatAlertsMentionEveryone/Here setters: adds or removes the ping in the mention text.</summary>
    public DiscordWebhookSettings WithMention(string ping, bool on)
    {
        bool has = Mention.Contains(ping, StringComparison.Ordinal);
        if (on && !has) return this with { Mention = string.Join(" ", new[] { Mention, ping }.Where(part => !string.IsNullOrWhiteSpace(part))).Trim() };
        if (!on && has) return this with { Mention = Mention.Replace(ping, "", StringComparison.Ordinal).Replace("  ", " ", StringComparison.Ordinal).Trim() };
        return this;
    }

    /// <summary>Reference SendDiscordWebhookAsync content: "{mention}\n" when set, then "**[{server}]** {message}".</summary>
    public static string Content(string mention, string? serverName, string message) =>
        (string.IsNullOrWhiteSpace(mention) ? "" : mention + "\n") + $"**[{serverName ?? "Rust Server"}]** {message}";
}

/// <summary>The selected profile's webhook settings in profiles.json; writes change only those fields.</summary>
public interface IDiscordWebhookStore
{
    Task<DiscordWebhookSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct);
    Task SetAsync(string profilesPath, string profileKey, DiscordWebhookSettings settings, CancellationToken ct);
}

/// <summary>Posts {content, tts} to a Discord webhook URL; throws when Discord rejects it.</summary>
public interface IDiscordWebhookSender
{
    Task SendAsync(string url, string content, bool tts, CancellationToken ct);
}

/// <summary>Reference IsCloudConnected: whether a Rust+ Desktop cloud session is stored and not expired.</summary>
public interface ICloudSignInState
{
    Task<bool> IsSignedInAsync(CancellationToken ct);
}
