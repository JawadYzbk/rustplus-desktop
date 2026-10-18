namespace RustPlusDesk.Core.Settings;

/// <summary>
/// Reference streamer mode (MapAbbreviateNames, "Abbreviates Steam Names and IDs to protect privacy during streams"):
/// names keep their first character, Steam IDs their first three digits, server names are replaced.
/// </summary>
public static class StreamerMode
{
    public const string HiddenServer = "SERVER HIDDEN";

    /// <summary>Reference GetDisplayPlayerName / TeamMemberVM.DisplayName.</summary>
    public static string Name(string name, bool on) => !on || string.IsNullOrWhiteSpace(name) ? name : name[..1] + "...";

    /// <summary>Reference TeamMemberVM.DisplaySteamId.</summary>
    public static string SteamId(string id, bool on) => !on ? id : id.Length <= 3 ? id + "..." : id[..3] + "...";
}
