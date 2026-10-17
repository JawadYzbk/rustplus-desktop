using System.Text.Json;
using RustPlusDesk.Core.Raid;

namespace RustPlusDesk.Infrastructure.Raid;

/// <summary>Loads and validates the reference raid-data.json bundled with the app; no network access.</summary>
public sealed class BundledRaidCatalog : IRaidCatalog
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private static readonly Lazy<RaidDataSet> Data = new(Read);
    public RaidDataSet Load() => Data.Value;

    private static RaidDataSet Read()
    {
        using var stream = typeof(BundledRaidCatalog).Assembly.GetManifestResourceStream("RustPlusDesk.Infrastructure.RaidData.json")
            ?? throw new InvalidDataException("The bundled raid dataset is missing.");
        return Parse(stream);
    }

    public static RaidDataSet Parse(Stream stream)
    {
        var data = JsonSerializer.Deserialize<RaidDataSet>(stream, JsonOptions) ?? throw new InvalidDataException("Raid data is empty.");
        RemoveUnsupportedTargets(data);
        Validate(data);
        return data;
    }

    /// <summary>Boat building blocks are not raidable structures in the reference calculator.</summary>
    private static void RemoveUnsupportedTargets(RaidDataSet data)
    {
        var targetIds = data.Targets.Where(target => target.PrefabName.Contains("/building boat/", StringComparison.OrdinalIgnoreCase))
            .Select(target => target.TargetId).ToHashSet();
        if (targetIds.Count == 0) return;
        data.Targets.RemoveAll(target => targetIds.Contains(target.TargetId));
        foreach (var values in data.DamagePerHit.Values) foreach (long id in targetIds) values.Remove(id);
        foreach (var values in data.Hits.Values) foreach (long id in targetIds) values.Remove(id);
    }

    public static void Validate(RaidDataSet data)
    {
        if (data.SchemaVersion != 1)
            throw new InvalidDataException($"Unsupported raid data schema version {data.SchemaVersion}.");
        if (data.Sources.Count == 0 || data.Targets.Count == 0 || data.Hits.Count == 0 || data.DamagePerHit.Count == 0)
            throw new InvalidDataException("Raid data must contain sources, targets, hit counts, and damage values.");
        var sourceIds = new HashSet<long>();
        foreach (var source in data.Sources)
        {
            if (source.SourceId <= 0 || !sourceIds.Add(source.SourceId) || string.IsNullOrWhiteSpace(source.DisplayName))
                throw new InvalidDataException("Raid data contains an invalid or duplicate source.");
            if (!double.IsFinite(source.RawDamage) || source.RawDamage < 0 ||
                source.DamageTypes.Any(value => !double.IsFinite(value.Value) || value.Value < 0) ||
                source.CraftCost?.Any(cost => string.IsNullOrWhiteSpace(cost.Shortname) || !double.IsFinite(cost.Amount) || cost.Amount < 0) == true)
                throw new InvalidDataException($"Raid source '{source.DisplayName}' contains malformed numeric data.");
        }
        var targetIds = new HashSet<long>();
        foreach (var target in data.Targets)
            if (target.TargetId <= 0 || !targetIds.Add(target.TargetId) || string.IsNullOrWhiteSpace(target.DisplayName) ||
                !double.IsFinite(target.StartHealth) || target.StartHealth <= 0)
                throw new InvalidDataException("Raid data contains an invalid or duplicate target.");
        ValidateMatrix(data.DamagePerHit, sourceIds, targetIds, value => double.IsFinite(value) && value > 0, "damage");
        ValidateMatrix(data.Hits, sourceIds, targetIds, value => value > 0, "hit count");
    }

    private static void ValidateMatrix<T>(Dictionary<long, Dictionary<long, T>> matrix, HashSet<long> sourceIds, HashSet<long> targetIds,
        Func<T, bool> valid, string name)
    {
        foreach ((long sourceId, var values) in matrix)
        {
            if (!sourceIds.Contains(sourceId)) throw new InvalidDataException($"Raid matrix references unknown source {sourceId}.");
            foreach ((long targetId, T value) in values)
                if (!targetIds.Contains(targetId) || !valid(value))
                    throw new InvalidDataException($"Raid matrix contains an invalid {name} for target {targetId}.");
        }
    }
}
