namespace RustPlusDesk.Core.Map;

/// <summary>
/// The per-profile Custom HD map image URL (reference ServerProfile.CustomMapUrl in profiles.json). Reads and writes adopt
/// the latest profiles document and change only that field of the profile with the given key.
/// </summary>
public interface ICustomMapUrlStore
{
    Task<string?> GetAsync(string profilesPath, string profileKey, CancellationToken ct);
    /// <summary>Stores the URL (null clears it). Throws <see cref="Servers.ProfileChangedException"/> when the profile is gone.</summary>
    Task SetAsync(string profilesPath, string profileKey, string? url, CancellationToken ct);
}

/// <summary>Downloads a custom HD map image (reference DownloadCustomImageFromUrlAsync); null on any failure.</summary>
public interface ICustomMapDownloader
{
    Task<byte[]?> DownloadAsync(string url, CancellationToken ct);
}

/// <summary>Reference BtnApplyCustomMapUrl_Click rules and texts.</summary>
public static class CustomMapUrl
{
    public const string InvalidMessage = "Please enter a valid absolute HTTP or HTTPS URL (e.g. https://example.com/rust_map.png).";
    public const string ActiveStatus = "Custom HD Map active. Map image is cached locally and will automatically reset on server wipe.";
    public const string NoneStatus = "No custom map URL set. Standard server map image is currently in use.";
    public const string UpdatedStatus = "Custom HD Map URL updated. Reloading map image...";
    public const string ClearedStatus = "Custom HD Map URL cleared. Reverting to standard server map...";
    public const string NoServerStatus = "No active server profile selected.";

    /// <summary>
    /// The text is trimmed; empty means "no custom map" (<paramref name="url"/> null). Anything else must be a well-formed
    /// absolute URI, as in the reference.
    /// </summary>
    public static bool TryNormalize(string? text, out string? url)
    {
        string trimmed = text?.Trim() ?? "";
        url = trimmed.Length == 0 ? null : trimmed;
        return url is null || Uri.IsWellFormedUriString(url, UriKind.Absolute);
    }
}
