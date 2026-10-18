using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Core.Deaths;

/// <summary>
/// Reference DeathTracker: turns successive team-info snapshots into death events. A death is detected when a member's
/// DeathTime advances (authoritative), or, when Rust+ reports no DeathTime, on an alive→dead transition stamped with the
/// current time. The first snapshot only establishes a baseline: deaths already present when the app connects are
/// history. Per-member de-duplication on the last recorded DeathTime stops the same death repeating while dead.
/// </summary>
public sealed class DeathTracker
{
    private readonly Dictionary<ulong, long> lastDeathTime = [];
    private readonly Dictionary<ulong, long> spawnObserved = [];
    private readonly Dictionary<ulong, bool> wasDeadById = [];
    private bool baselineEstablished;

    /// <summary>Reset when switching servers so one server's state can't leak into another.</summary>
    public void Reset()
    {
        lastDeathTime.Clear();
        spawnObserved.Clear();
        wasDeadById.Clear();
        baselineEstablished = false;
    }

    public IReadOnlyList<DeathRecord> Observe(IReadOnlyList<ServerTeamMember> members, DeathLocationClassifier classifier, long now)
    {
        var deaths = new List<DeathRecord>();
        foreach (var member in members)
        {
            if (!ulong.TryParse(member.SteamId, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out ulong id) || id == 0)
                continue;
            // The reference reads "dead" from IsAlive; an unknown state counts as alive.
            bool dead = member.IsAlive == false;
            bool wasDead = wasDeadById.TryGetValue(id, out bool previous) && previous;

            // A dead→alive edge is a respawn: remember when, so the next death's survival time is known without SpawnTime.
            if (wasDead && !dead) spawnObserved[id] = member.SpawnTime is > 0 ? member.SpawnTime.Value : now;
            wasDeadById[id] = dead;

            bool isNewDeath;
            long deathTime;
            if (member.DeathTime is > 0)
            {
                long last = lastDeathTime.TryGetValue(id, out long seen) ? seen : 0;
                isNewDeath = member.DeathTime.Value > last;
                deathTime = member.DeathTime.Value;
            }
            else
            {
                isNewDeath = dead && !wasDead;
                deathTime = now;
            }
            if (!isNewDeath) continue;
            lastDeathTime[id] = deathTime;
            if (!baselineEstablished) continue;

            long? spawnTime = member.SpawnTime is > 0 ? member.SpawnTime.Value : spawnObserved.TryGetValue(id, out long observed) ? observed : null;
            var (type, name) = classifier.Classify(member.X, member.Y);
            deaths.Add(new DeathRecord(id, member.Name, deathTime, spawnTime, member.X, member.Y, classifier.Grid(member.X, member.Y), type, name));
        }
        baselineEstablished = true;
        return deaths;
    }
}
