using System.Text.Json;
using RustPlusDesk.Core.Devices;

namespace RustPlusDesk.Infrastructure.Items;

/// <summary>Uses the same packaged item names as the reference app, without UI or network dependencies.</summary>
public static class BundledItemCatalog
{
    private static readonly Lazy<IReadOnlyDictionary<int, DeviceIcon>> Entries = new(Load);
    public static IReadOnlyDictionary<int, DeviceIcon> Icons => Entries.Value;

    public static string GetName(int id) => Icons.TryGetValue(id, out var item) ? item.Name : $"Item #{id}";

    private static IReadOnlyDictionary<int, DeviceIcon> Load()
    {
        using var stream = typeof(BundledItemCatalog).Assembly.GetManifestResourceStream("RustPlusDesk.Infrastructure.Items.json")
            ?? throw new InvalidDataException("The packaged item catalog is missing.");
        using var document = JsonDocument.Parse(stream);
        var names = new Dictionary<int, DeviceIcon>();
        foreach (var item in document.RootElement.EnumerateArray())
        {
            int id = item.GetProperty("id").GetInt32();
            string? name = item.GetProperty("displayName").GetString();
            string? shortName = item.GetProperty("shortName").GetString();
            if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(shortName)) names[id] = new(id, shortName, name);
        }
        if (names.Count == 0) throw new InvalidDataException("The packaged item catalog is empty.");
        return names;
    }
}
