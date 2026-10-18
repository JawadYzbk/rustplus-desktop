using System.Text.RegularExpressions;

namespace RustPlusDesk.Core.Connections;

/// <summary>Loads a player's Steam avatar image (reference AvatarLoader); null when the profile is private or offline.</summary>
public interface ISteamAvatarSource
{
    Task<byte[]?> LoadAsync(ulong steamId, CancellationToken ct);
}

/// <summary>Reference AvatarLoader parsing of steamcommunity.com/profiles/{id}?xml=1.</summary>
public static partial class SteamAvatarXml
{
    public const string ProfileUrlFormat = "https://steamcommunity.com/profiles/{0}?xml=1";
    /// <summary>Reference: retry a failed avatar after 30 seconds.</summary>
    public static readonly TimeSpan RetryInterval = TimeSpan.FromSeconds(30);

    /// <summary>The full-size avatar URL, else the medium one, else null.</summary>
    public static string? AvatarUrl(string xml)
    {
        var full = Full().Match(xml);
        if (full.Success && !string.IsNullOrWhiteSpace(full.Groups[1].Value)) return full.Groups[1].Value;
        var medium = Medium().Match(xml);
        return medium.Success && !string.IsNullOrWhiteSpace(medium.Groups[1].Value) ? medium.Groups[1].Value : null;
    }

    [GeneratedRegex(@"<avatarFull><!\[CDATA\[(.*?)\]\]></avatarFull>", RegexOptions.IgnoreCase)]
    private static partial Regex Full();
    [GeneratedRegex(@"<avatarMedium><!\[CDATA\[(.*?)\]\]></avatarMedium>", RegexOptions.IgnoreCase)]
    private static partial Regex Medium();
}
