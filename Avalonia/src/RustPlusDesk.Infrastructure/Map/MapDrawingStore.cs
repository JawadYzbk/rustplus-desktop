using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Map;

/// <summary>
/// Reference OverlayDataModule: the drawings the user made on a map, saved in the app's data folder under
/// <c>Overlays/&lt;server&gt;/&lt;steam id&gt;.json</c> so they are found again for the same server and the same player.
///
/// The file is written atomically and read defensively: a file that cannot be read leaves the map without drawings
/// rather than stopping the app.
/// </summary>
public sealed class MapDrawingStore(string root)
{
    public static string DefaultRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "Overlays");

    /// <summary>Reference GetOverlayJsonPath: one file per server and per Steam id.</summary>
    public string PathFor(string serverKey, string steamId) =>
        Path.Combine(root, Safe(serverKey), $"{Safe(steamId)}.json");

    /// <summary>
    /// A server key can hold a host and a port, and a colon is not allowed in a folder name on Windows, so the
    /// characters that cannot be used are replaced rather than letting the write fail.
    /// </summary>
    private static string Safe(string name)
    {
        var builder = new System.Text.StringBuilder(name.Length);
        foreach (char character in name)
            builder.Append(char.IsLetterOrDigit(character) || character is '-' or '_' or '.' ? character : '_');
        return builder.ToString();
    }

    public MapOverlayData Load(string serverKey, string steamId)
    {
        string path = PathFor(serverKey, steamId);
        try
        {
            if (!File.Exists(path)) return MapOverlayData.Empty;
            if (JsonNode.Parse(File.ReadAllBytes(path)) is not JsonObject document) return MapOverlayData.Empty;
            return Read(document);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            return MapOverlayData.Empty;
        }
    }

    public void Save(string serverKey, string steamId, MapOverlayData data)
    {
        string path = PathFor(serverKey, steamId);
        try
        {
            string? folder = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
            var document = new JsonObject
            {
                ["Strokes"] = new JsonArray(data.Strokes.Select(stroke => (JsonNode)new JsonObject
                {
                    ["Points"] = new JsonArray(stroke.Points.Select(point => (JsonNode)new JsonObject
                    {
                        ["X"] = point.X,
                        ["Y"] = point.Y
                    }).ToArray()),
                    ["Color"] = stroke.Color,
                    ["Thickness"] = stroke.Thickness,
                    ["GroupId"] = stroke.GroupId
                }).ToArray()),
                ["Icons"] = new JsonArray(data.Icons.Select(icon => (JsonNode)new JsonObject
                {
                    ["IconPath"] = icon.IconPath,
                    ["X"] = icon.X,
                    ["Y"] = icon.Y,
                    ["Size"] = icon.Size,
                    ["Label"] = icon.Label
                }).ToArray()),
                ["Texts"] = new JsonArray(data.Texts.Select(text => (JsonNode)new JsonObject
                {
                    ["X"] = text.X,
                    ["Y"] = text.Y,
                    ["Text"] = text.Text,
                    ["Color"] = text.Color,
                    ["Size"] = text.Size
                }).ToArray()),
                ["Routes"] = new JsonArray(data.Routes.Select(route => (JsonNode)new JsonObject
                {
                    ["Id"] = route.Id,
                    ["Name"] = route.Name,
                    ["Color"] = route.Color,
                    ["Thickness"] = route.Thickness,
                    ["Points"] = new JsonArray(route.Points.Select(point => (JsonNode)new JsonObject
                    {
                        ["X"] = point.X,
                        ["Y"] = point.Y
                    }).ToArray()),
                    ["Closed"] = route.Closed,
                    ["Visible"] = route.Visible,
                    ["SourceId"] = route.SourceId
                }).ToArray()),
                ["WipeTimeUnix"] = data.WipeTimeUnix
            };
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllText(temporary, document.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
            File.Move(temporary, path, overwrite: true);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            // A failed write must never stop the app; the drawing is still on the map for this session.
        }
    }

    private static MapOverlayData Read(JsonObject document) => new()
    {
        Strokes = Strokes(document),
        Icons = Icons(document),
        Texts = Texts(document),
        Routes = Routes(document),
        WipeTimeUnix = document["WipeTimeUnix"] is JsonValue value && value.TryGetValue(out long wipe) ? wipe : 0
    };

    private static IReadOnlyList<SavedStroke> Strokes(JsonObject document)
    {
        if (document["Strokes"] is not JsonArray rows) return [];
        var strokes = new List<SavedStroke>();
        foreach (var row in rows)
        {
            if (row is not JsonObject stroke) continue;
            strokes.Add(new SavedStroke(Points(stroke), Text(stroke, "Color") ?? "#FF0000",
                Number(stroke, "Thickness") ?? 2.0, Text(stroke, "GroupId")));
        }
        return strokes;
    }

    private static IReadOnlyList<SavedIcon> Icons(JsonObject document)
    {
        if (document["Icons"] is not JsonArray rows) return [];
        var icons = new List<SavedIcon>();
        foreach (var row in rows)
        {
            if (row is not JsonObject icon) continue;
            string path = Text(icon, "IconPath") ?? "";
            if (path.Length == 0) continue;
            icons.Add(new SavedIcon(path, Number(icon, "X") ?? 0, Number(icon, "Y") ?? 0,
                Number(icon, "Size") ?? 24, Text(icon, "Label")));
        }
        return icons;
    }

    private static IReadOnlyList<SavedText> Texts(JsonObject document)
    {
        if (document["Texts"] is not JsonArray rows) return [];
        var texts = new List<SavedText>();
        foreach (var row in rows)
        {
            if (row is not JsonObject text) continue;
            string value = Text(text, "Text") ?? "";
            if (value.Length == 0) continue;
            texts.Add(new SavedText(Number(text, "X") ?? 0, Number(text, "Y") ?? 0, value,
                Text(text, "Color") ?? "#FFFFFFFF", Number(text, "Size") ?? 14));
        }
        return texts;
    }

    private static IReadOnlyList<SavedRoute> Routes(JsonObject document)
    {
        if (document["Routes"] is not JsonArray rows) return [];
        var routes = new List<SavedRoute>();
        foreach (var row in rows)
        {
            if (row is not JsonObject route) continue;
            string id = Text(route, "Id") ?? "";
            if (id.Length == 0) continue;
            routes.Add(new SavedRoute(id, Text(route, "Name") ?? "", Text(route, "Color") ?? "#FF3FD7FF",
                Number(route, "Thickness") ?? 3.0, Points(route),
                route["Closed"] is JsonValue closed && closed.TryGetValue(out bool isClosed) && isClosed,
                route["Visible"] is not JsonValue visible || !visible.TryGetValue(out bool hidden) || hidden,
                Text(route, "SourceId")));
        }
        return routes;
    }

    private static IReadOnlyList<MapPoint> Points(JsonObject node)
    {
        if (node["Points"] is not JsonArray rows) return [];
        var points = new List<MapPoint>();
        foreach (var row in rows)
        {
            if (row is not JsonObject point) continue;
            points.Add(new MapPoint(Number(point, "X") ?? 0, Number(point, "Y") ?? 0));
        }
        return points;
    }

    private static string? Text(JsonObject node, string name) =>
        node[name] is JsonValue value && value.TryGetValue(out string? text) ? text : null;

    private static double? Number(JsonObject node, string name) =>
        node[name] is JsonValue value && value.TryGetValue(out double number) ? number : null;
}
