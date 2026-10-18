using System.Text;
using System.Text.Json.Serialization;

namespace RustPlusDesk.Core.Servers;

/// <summary>A credential-free description of a saved server.</summary>
public sealed record ServerSummary(string Name, string Host, int Port)
{
    public string ProfileKey { get; init; } = "";
    /// <summary>
    /// Reference ServerProfile.MatchKey ("host:port|steamId64"), only for the LastSelectedServerKey setting the reference
    /// shares. It carries the Steam ID, so it is kept out of serialization and <see cref="ToString"/>.
    /// </summary>
    [JsonIgnore] public string MatchKey { get; init; } = "";
    public string SourceRevision { get; init; } = "";
    public IReadOnlyList<Devices.PairedDevice> Devices { get; init; } = [];

    private bool PrintMembers(StringBuilder builder)
    {
        builder.Append($"Name = {Name}, Host = {Host}, Port = {Port}, ProfileKey = {ProfileKey}, SourceRevision = {SourceRevision}, Devices = {Devices.Count}");
        return true;
    }
}
