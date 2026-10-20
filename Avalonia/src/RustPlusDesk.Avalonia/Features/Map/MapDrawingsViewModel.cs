using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Infrastructure.Map;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>One named route, as the routes list shows it (reference SavedRoute).</summary>
public sealed partial class RouteRowViewModel(SavedRoute route) : ObservableObject
{
    public string Id { get; } = route.Id;
    public string Name { get; private set; } = route.Name;
    public string Color { get; private set; } = route.Color;
    public int PointCount { get; private set; } = route.Points.Count;

    /// <summary>Hidden routes stay in the list and off the map (reference SavedRoute.Visible).</summary>
    [ObservableProperty] private bool isVisible = route.Visible;

    public SavedRoute ToRoute() => route with { Visible = IsVisible };
}

/// <summary>One drawing tool, as the tool picker shows it.</summary>
public sealed record MapToolChoice(string Label, MapToolMode Tool);

/// <summary>
/// Reference the map's overlay: the strokes, the routes, the icons and the texts the user drew, saved per server.
///
/// The drawings are the user's own work, so they are kept in the app's data folder under
/// <c>Overlays/&lt;server&gt;/&lt;steam id&gt;.json</c> and found again for the same server.
/// </summary>
public sealed partial class MapDrawingsViewModel(MapDrawingStore store) : ObservableObject
{
    /// <summary>The tools the map offers, in the reference's order.</summary>
    public static IReadOnlyList<MapToolChoice> ToolChoices { get; } =
    [
        new("Select (pan the map)", MapToolMode.None),
        new("Line", MapToolMode.Line),
        new("Arrow", MapToolMode.Arrow),
        new("Box", MapToolMode.Box),
        new("Circle", MapToolMode.Circle),
        new("Route", MapToolMode.Route)
    ];

    /// <summary>The server the drawings belong to (the reference's server key).</summary>
    public Func<string>? ServerKey { get; set; }

    /// <summary>The Steam id of the player who drew them.</summary>
    public Func<string?>? SteamId { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Tool))]
    private MapToolChoice toolChoice = ToolChoices[0];

    /// <summary>The tool the user draws with (reference OverlayToolMode).</summary>
    public MapToolMode Tool => ToolChoice.Tool;

    [ObservableProperty] private string color = "#FFFF3B30";
    [ObservableProperty] private double thickness = 2;
    [ObservableProperty] private MapOverlayData overlay = MapOverlayData.Empty;
    [ObservableProperty] private string routeName = "";
    [ObservableProperty] private string status = "";
    [ObservableProperty] private bool isDrawing;

    public ObservableCollection<RouteRowViewModel> Routes { get; } = [];

    /// <summary>The points of the route being drawn, in world coordinates.</summary>
    private List<MapPoint> pending = [];

    /// <summary>Reference LoadLocalOverlay: the drawings are read for the selected server.</summary>
    public void Load()
    {
        string server = ServerKey?.Invoke() ?? "";
        string steam = SteamId?.Invoke() ?? "";
        if (server.Length == 0 || steam.Length == 0) { Overlay = MapOverlayData.Empty; RefreshRoutes(); return; }
        Overlay = store.Load(server, steam);
        RefreshRoutes();
    }

    /// <summary>Reference SaveLocalOverlay: the drawings are written back.</summary>
    public void Save()
    {
        string server = ServerKey?.Invoke() ?? "";
        string steam = SteamId?.Invoke() ?? "";
        if (server.Length == 0 || steam.Length == 0) return;
        store.Save(server, steam, Overlay);
    }

    /// <summary>Reference the shape tools: a drag becomes a stroke.</summary>
    public void DrawShape(MapPoint from, MapPoint to)
    {
        var points = MapShapes.BuildShapePoints(Tool, from, to, false, Thickness);
        if (points.Count < 2) return;
        Overlay = Overlay with { Strokes = [.. Overlay.Strokes, new SavedStroke(points, Color, Thickness)] };
        Save();
    }

    /// <summary>Reference the route tool: a click adds a point to the route being drawn.</summary>
    public void PlacePoint(MapPoint point)
    {
        if (Tool != MapToolMode.Route) return;
        pending.Add(point);
        IsDrawing = pending.Count > 0;
    }

    /// <summary>Reference the route list: the route is named and kept.</summary>
    [RelayCommand]
    public void FinishRoute()
    {
        if (pending.Count < 2) { Status = "A route needs at least two points."; return; }
        string name = string.IsNullOrWhiteSpace(RouteName) ? $"Route {Routes.Count + 1}" : RouteName.Trim();
        var route = new SavedRoute(Guid.NewGuid().ToString("N"), name, Color, Thickness,
            MapShapes.SimplifyPoints(pending));
        Overlay = Overlay with { Routes = [.. Overlay.Routes, route] };
        pending = [];
        IsDrawing = false;
        RouteName = "";
        RefreshRoutes();
        Save();
    }

    /// <summary>Reference CancelActiveOverlayDraw: the route being drawn is thrown away.</summary>
    [RelayCommand]
    public void CancelRoute()
    {
        pending = [];
        IsDrawing = false;
        Status = "";
    }

    /// <summary>Reference the undo: the last drawing is taken back.</summary>
    [RelayCommand]
    public void Undo()
    {
        if (Overlay.Strokes.Count == 0) return;
        Overlay = Overlay with { Strokes = Overlay.Strokes.Take(Overlay.Strokes.Count - 1).ToArray() };
        Save();
    }

    /// <summary>Reference the trash: every drawing on this map is removed.</summary>
    [RelayCommand]
    public void ClearAll()
    {
        Overlay = MapOverlayData.Empty;
        pending = [];
        IsDrawing = false;
        RefreshRoutes();
        Save();
    }

    /// <summary>Reference the route list: a hidden route stays in the list and off the map.</summary>
    [RelayCommand]
    public void ToggleVisibility(RouteRowViewModel? row)
    {
        if (row is null) return;
        Overlay = Overlay with
        {
            Routes = Overlay.Routes.Select(route => route.Id == row.Id ? row.ToRoute() : route).ToArray()
        };
        Save();
    }

    /// <summary>Reference the route delete: the route leaves the list and the map.</summary>
    [RelayCommand]
    public void DeleteRoute(RouteRowViewModel? row)
    {
        if (row is null) return;
        Overlay = Overlay with { Routes = Overlay.Routes.Where(route => route.Id != row.Id).ToArray() };
        RefreshRoutes();
        Save();
    }

    private void RefreshRoutes()
    {
        Routes.Clear();
        foreach (var route in Overlay.Routes) Routes.Add(new RouteRowViewModel(route));
    }
}
