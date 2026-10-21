using System.Globalization;
using RustPlusDesk.Avalonia.Features.Connections;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MainWindow.StartSessionTracking: the one reading a second the session tracker counts from.
///
/// Everything the tracker counts comes from here - the connection, the server, the player's own position and idleness,
/// the population of the server and the team with each member's state - so the tracker itself holds no game knowledge at
/// all.
/// </summary>
public sealed class DockSessionSource(
    ConnectionViewModel connection,
    Func<string?> serverKey,
    Func<ulong?> ownSteamId)
{
    /// <summary>Reference SessionTracker.Source: the reading of right now, or null when there is nothing to count.</summary>
    public SessionSnapshot? Read()
    {
        if (!connection.IsConnected) return null;

        string? key = serverKey();
        if (string.IsNullOrEmpty(key)) return null;

        ulong mySteamId = ownSteamId() ?? 0;
        var members = connection.Team.Members;
        var me = members.FirstOrDefault(row => SteamId(row.Member.SteamId) == mySteamId && mySteamId != 0);

        var team = new List<(ulong SteamId, bool IsDead)>(members.Count);
        foreach (var row in members)
        {
            ulong id = SteamId(row.Member.SteamId);
            if (id == 0) continue;
            team.Add((id, row.IsDead));
        }

        return new SessionSnapshot(
            Connected: true,
            ServerKey: key,
            MySteamId: mySteamId,
            MyX: me?.Member.X,
            MyY: me?.Member.Y,
            MyAfk: me?.IsAfk == true,
            Population: Population(),
            Team: team);
    }

    /// <summary>Reference MainWindow.ParsePopulation: "128/200" as 128.</summary>
    private int Population()
    {
        string text = connection.Population;
        if (string.IsNullOrWhiteSpace(text)) return 0;
        string head = text.Split('/')[0].Trim();
        return int.TryParse(head, NumberStyles.Integer, CultureInfo.InvariantCulture, out int players) ? players : 0;
    }

    /// <summary>The member's Steam ID, or zero when it is not one.</summary>
    private static ulong SteamId(string? id) =>
        ulong.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out ulong steamId) ? steamId : 0;
}
