using System.Globalization;

namespace RustPlusDesk.Core.Devices;

/// <summary>
/// Reference ServerProfile.DeviceAutomationRules: one rule that puts a paired switch into a state when its condition
/// holds. The rule is saved with the profile, so it survives a restart and travels with the server.
/// </summary>
public sealed record DeviceAutomationRule
{
    public string Id { get; init; } = Guid.NewGuid().ToString();
    public string Name { get; init; } = "New Automation";
    public bool IsEnabled { get; init; }
    /// <summary>"PlayerProximity" or "GameTime" (reference DeviceAutomationRule.ConditionType).</summary>
    public string ConditionType { get; init; } = "PlayerProximity";
    /// <summary>
    /// How the players are matched: "AnyOnline", "AllOnline", "AnyOffline", "AllOffline", "Specific" or
    /// "SpecificOffline" (reference DeviceAutomationRule.PlayerMatchMode).
    /// </summary>
    public string PlayerMatchMode { get; init; } = "AnyOnline";
    public ulong SpecificPlayerSteamId { get; init; }
    /// <summary>The device whose captured position is the anchor of the proximity test.</summary>
    public uint LocationEntityId { get; init; }
    public double DistanceMeters { get; init; } = 250;
    public string StartTime { get; init; } = "20:00";
    public string EndTime { get; init; } = "08:00";
    /// <summary>The switch the rule drives.</summary>
    public uint TargetEntityId { get; init; }
    /// <summary>The state the target takes while the condition holds.</summary>
    public bool MatchedState { get; init; }
    /// <summary>The state the target takes while the condition does not hold; null leaves the switch alone.</summary>
    public bool? UnmatchedState { get; init; } = true;
    /// <summary>What the rule last wrote, so the card can show it (reference LastAppliedState).</summary>
    public bool? LastAppliedState { get; init; }
}

/// <summary>
/// Reference ServerProfile.IsDeviceAutomationActive and DeviceAutomationRules: the rules of one profile and the
/// switch that turns them on.
/// </summary>
public sealed record DeviceAutomationSettings(bool Enabled = false, IReadOnlyList<DeviceAutomationRule> Rules = null!)
{
    /// <summary>Shared empty rule list, so two settings records with no rules compare equal.</summary>
    public static IReadOnlyList<DeviceAutomationRule> NoRules { get; } = [];

    public static DeviceAutomationSettings Defaults { get; } = new(Rules: NoRules);

    public DeviceAutomationSettings() : this(false, NoRules) { }
}

/// <summary>
/// The selected profile's device-automation rules. Rules live in the profile, so they travel with the server and
/// survive a restart.
/// </summary>
public interface IDeviceAutomationStore
{
    Task<DeviceAutomationSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct);

    /// <summary>Throws <c>ProfileChangedException</c> when the profile no longer exists.</summary>
    Task SaveAsync(string profilesPath, string profileKey, DeviceAutomationSettings settings, CancellationToken ct);

    /// <summary>
    /// Reference CapturePairedDeviceLocationAsync: records where a device was seen, so a proximity rule can use it as
    /// its anchor. The position is written on the device's own row.
    /// </summary>
    Task CaptureLocationAsync(string profilesPath, string profileKey, uint entityId, double x, double y,
        ulong steamId, CancellationToken ct);
}

/// <summary>
/// Reference DeviceAutomationEvaluator: the two conditions a rule can test. The rules are pure so they can be
/// checked without a live server.
/// </summary>
public static class DeviceAutomationEvaluator
{
    /// <summary>One teammate as the proximity test sees them.</summary>
    public readonly record struct PlayerSnapshot(ulong SteamId, bool IsOnline, double? X, double? Y);

    /// <summary>
    /// Reference IsProximityMatch: a player near the anchor counts when they are online, have a position and are
    /// within the rule's distance. The offline modes look at the same selection without the position.
    /// </summary>
    public static bool IsProximityMatch(
        DeviceAutomationRule rule,
        double anchorX,
        double anchorY,
        IReadOnlyCollection<PlayerSnapshot> players)
    {
        var selected = rule.PlayerMatchMode.StartsWith("Specific", StringComparison.Ordinal)
            ? players.Where(player => player.SteamId == rule.SpecificPlayerSteamId).ToList()
            : players.ToList();

        if (rule.PlayerMatchMode == "AnyOffline") return selected.Any(player => !player.IsOnline);
        if (rule.PlayerMatchMode == "AllOffline") return selected.Count > 0 && selected.All(player => !player.IsOnline);
        if (rule.PlayerMatchMode == "SpecificOffline") return selected.Count == 1 && !selected[0].IsOnline;

        bool IsNear(PlayerSnapshot player) => player.IsOnline
            && player.X.HasValue
            && player.Y.HasValue
            && Math.Sqrt(Math.Pow(player.X.Value - anchorX, 2) + Math.Pow(player.Y.Value - anchorY, 2)) <= rule.DistanceMeters;

        return rule.PlayerMatchMode == "AllOnline"
            ? selected.Count > 0 && selected.All(IsNear)
            : selected.Any(IsNear);
    }

    /// <summary>Reference IsTimeMatch: the game clock is inside the rule's window.</summary>
    public static bool IsTimeMatch(DeviceAutomationRule rule, string currentTime) =>
        TryGetTimeMatch(rule, currentTime, out bool matched) && matched;

    /// <summary>
    /// Reference TryGetTimeMatch: the window may wrap around midnight. False means the times could not be read at
    /// all, which is not the same as "outside the window".
    /// </summary>
    public static bool TryGetTimeMatch(DeviceAutomationRule rule, string currentTime, out bool matched)
    {
        matched = false;
        if (!TryMinutes(rule.StartTime, out int start)
            || !TryMinutes(rule.EndTime, out int end)
            || !TryMinutes(currentTime, out int current))
            return false;

        matched = start == end || (start < end
            ? current >= start && current < end
            : current >= start || current < end);
        return true;
    }

    private static bool TryMinutes(string value, out int minutes)
    {
        minutes = 0;
        if (!TimeSpan.TryParse(value, CultureInfo.InvariantCulture, out var time)) return false;
        minutes = ((int)time.TotalMinutes % 1440 + 1440) % 1440;
        return true;
    }
}

/// <summary>What a rule wants a switch to become (reference the decision list in EvaluateDeviceAutomationAsync).</summary>
public sealed record AutomationDecision(DeviceAutomationRule Rule, uint TargetEntityId, bool State);

/// <summary>
/// Reference EvaluateDeviceAutomationAsync: the rules are read into a list of decisions before anything is written, so
/// two rules that disagree about the same switch are spotted before one of them wins.
/// </summary>
public static class DeviceAutomationRunner
{
    /// <summary>
    /// The decisions the rules ask for. A target two rules disagree about is dropped and reported, exactly as the
    /// reference refuses to act on a conflict.
    /// </summary>
    public static (IReadOnlyList<AutomationDecision> Decisions, IReadOnlyList<uint> Conflicts) Decide(
        IReadOnlyList<DeviceAutomationRule> rules,
        IReadOnlyList<PairedDevice> devices,
        IReadOnlyList<DeviceAutomationEvaluator.PlayerSnapshot> players,
        string serverTime)
    {
        var decisions = new List<AutomationDecision>();
        foreach (var rule in rules.Where(rule => rule.IsEnabled))
        {
            bool matched;
            if (rule.ConditionType == "PlayerProximity")
            {
                double x = 0, y = 0;
                if (!rule.PlayerMatchMode.EndsWith("Offline", StringComparison.Ordinal))
                {
                    var location = Find(devices, rule.LocationEntityId);
                    if (location?.PairedX is not double pairedX || location.PairedY is not double pairedY) continue;
                    x = pairedX;
                    y = pairedY;
                }
                matched = DeviceAutomationEvaluator.IsProximityMatch(rule, x, y, players);
            }
            else if (rule.ConditionType == "GameTime")
            {
                if (!DeviceAutomationEvaluator.TryGetTimeMatch(rule, serverTime, out matched)) continue;
            }
            else continue;

            var target = Find(devices, rule.TargetEntityId);
            if (target is null || target.Kind != DeviceKind.Switch) continue;
            // "None" for the unmatched case: the rule only acts while its condition holds.
            bool? state = matched ? rule.MatchedState : rule.UnmatchedState;
            if (state is not bool desired) continue;
            decisions.Add(new AutomationDecision(rule, target.EntityId, desired));
        }

        var conflicts = new List<uint>();
        var agreed = new List<AutomationDecision>();
        foreach (var group in decisions.GroupBy(decision => decision.TargetEntityId))
        {
            var states = group.Select(decision => decision.State).Distinct().ToList();
            if (states.Count != 1) { conflicts.Add(group.Key); continue; }
            agreed.Add(group.First());
        }
        return (agreed, conflicts);
    }

    private static PairedDevice? Find(IReadOnlyList<PairedDevice> devices, uint entityId) =>
        devices.SelectMany(device => device.DescendantsAndSelf()).FirstOrDefault(device => device.EntityId == entityId);
}
