using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Win32;
using RustPlusDesk.Services.PlayerWipeTracker;

namespace RustPlusDesk.Views.Windows;

public partial class PlayerWipeTrackerWindow : Window
{
    private readonly PlayerWipeTrackerService _tracker;
    private readonly ulong _ownSteamId;
    private readonly ImageSource? _wipeMapImage;
    private readonly int _worldSize;
    private readonly Rect _worldRectPixels;
    private readonly DispatcherTimer _replayTimer;
    private IReadOnlyList<TrackerPoint> _points = Array.Empty<TrackerPoint>();
    private double _replaySpeed = 1;
    private int _replayIndex;
    private bool _refreshing;

    public PlayerWipeTrackerWindow(
        PlayerWipeTrackerService tracker,
        ulong ownSteamId,
        ImageSource? wipeMapImage,
        int worldSize,
        Rect worldRectPixels)
    {
        _tracker = tracker ?? throw new ArgumentNullException(nameof(tracker));
        _ownSteamId = ownSteamId;
        _wipeMapImage = wipeMapImage;
        _worldSize = worldSize;
        _worldRectPixels = worldRectPixels;
        _replayTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
        _replayTimer.Tick += ReplayTimer_Tick;

        InitializeComponent();
        ReplayMapImage.Source = _wipeMapImage;
        HeatmapMapImage.Source = _wipeMapImage;
        Refresh();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e) => Refresh();

    private void Canvas_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        RenderReplay();
        RenderHeatmap();
        RenderComparison();
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => Refresh();

    private void PlayerSelector_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (!_refreshing)
            Refresh();
    }

    private void Refresh()
    {
        var capabilities = _tracker.Capabilities;
        PremiumBanner.Text = capabilities.CanUseAdvancedViews
            ? $"{capabilities.PlanCode.ToUpperInvariant()} FIELD CONSOLE · Cloud backup is opt-in. Route replay, heatmap, comparison, export, and restore are enabled."
            : "FREE FIELD CONSOLE · Self-only local history is available. Premium route replay, heatmap, comparison, export, and cloud restore are locked.";

        ReplayTab.IsEnabled = capabilities.CanUseRouteReplay;
        HeatmapTab.IsEnabled = capabilities.CanUseAdvancedViews;
        CompareTab.IsEnabled = capabilities.CanUseAdvancedViews;
        ExportTab.IsEnabled = capabilities.CanExport;
        CloudTab.IsEnabled = capabilities.CanUseCloudSync;
        RestoreCloudButton.IsEnabled = capabilities.CanUseCloudSync;
        ExportJsonButton.IsEnabled = capabilities.CanExport;
        ExportCsvButton.IsEnabled = capabilities.CanExport;

        var selected = SelectedPlayerId(PlayerSelector);
        var items = BuildPlayerItems();
        _refreshing = true;
        try
        {
            SetPlayerItems(PlayerSelector, items, selected ?? _ownSteamId);
            SetPlayerItems(CompareASelector, items, selected ?? _ownSteamId);
            SetPlayerItems(CompareBSelector, items, items.FirstOrDefault(item => item.SteamId != (selected ?? _ownSteamId))?.SteamId ?? selected ?? _ownSteamId);
        }
        finally
        {
            _refreshing = false;
        }

        RefreshSelectedPlayer();
    }

    private void RefreshSelectedPlayer()
    {
        var steamId = SelectedPlayerId(PlayerSelector) ?? _ownSteamId;
        var observations = _tracker.GetObservations(steamId);
        _points = ToPoints(observations);
        var summary = _tracker.GetSummary(steamId);

        CoverageText.Text = Format(summary.Coverage);
        UnknownText.Text = Format(summary.Unknown);
        DistanceText.Text = $"{summary.EstimatedDistance:N0}";
        DeathsText.Text = summary.Deaths.ToString();
        PointCountText.Text = _points.Count.ToString();
        StorageText.Text = FormatBytes(_tracker.StorageBytes);
        TimelineList.ItemsSource = _tracker.GetSegments(steamId)
            .OrderByDescending(segment => segment.StartUtc)
            .Select(segment => new
            {
                Start = segment.StartUtc.ToLocalTime().ToString("g"),
                End = segment.EndUtc.ToLocalTime().ToString("g"),
                State = segment.State.ToString(),
                Location = segment.LocationName ?? segment.LocationType.ToString(),
            }).ToArray();
        VisitsList.ItemsSource = summary.MonumentVisits
            .OrderByDescending(visit => visit.StartUtc)
            .Select(visit => new { visit.Name, Duration = Format(visit.EstimatedDuration) })
            .ToArray();

        _replayTimer.Stop();
        _replayIndex = Math.Max(0, _points.Count - 1);
        ReplayPlayButton.Content = "Play";
        ReplayProgress.Maximum = Math.Max(0, _points.Count - 1);
        ReplayProgress.Value = _replayIndex;
        ReplayProgress.IsEnabled = _points.Count > 1 && _tracker.Capabilities.CanUseRouteReplay;
        RenderReplay();
        RenderHeatmap();
        RenderComparison();
    }

    private List<PlayerItem> BuildPlayerItems()
    {
        var ids = new HashSet<ulong>(_tracker.TrackedPlayers) { _ownSteamId };
        return ids.Where(id => id != 0)
            .Select(id => new PlayerItem(id, _tracker.GetPlayerName(id)))
            .OrderBy(item => item.SteamId != _ownSteamId)
            .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void SetPlayerItems(ComboBox selector, IReadOnlyList<PlayerItem> items, ulong steamId)
    {
        selector.ItemsSource = items.ToArray();
        selector.SelectedItem = items.FirstOrDefault(item => item.SteamId == steamId) ?? items.FirstOrDefault();
    }

    private static ulong? SelectedPlayerId(Selector selector)
        => (selector.SelectedItem as PlayerItem)?.SteamId;

    private void ReplayPlay_Click(object sender, RoutedEventArgs e)
    {
        if (_points.Count < 2 || !_tracker.Capabilities.CanUseRouteReplay)
            return;
        if (_replayIndex >= _points.Count - 1)
            _replayIndex = 0;
        if (_replayTimer.IsEnabled)
        {
            _replayTimer.Stop();
            ReplayPlayButton.Content = "Play";
        }
        else
        {
            _replayTimer.Start();
            ReplayPlayButton.Content = "Pause";
        }
        RenderReplay();
    }

    private void ReplayTimer_Tick(object? sender, EventArgs e)
    {
        _replayIndex = Math.Min(_points.Count - 1, _replayIndex + Math.Max(1, (int)Math.Round(_replaySpeed)));
        ReplayProgress.Value = _replayIndex;
        if (_replayIndex >= _points.Count - 1)
        {
            _replayTimer.Stop();
            ReplayPlayButton.Content = "Play";
        }
        RenderReplay();
    }

    private void ReplayProgress_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        _replayIndex = Math.Clamp((int)Math.Round(e.NewValue), 0, Math.Max(0, _points.Count - 1));
        RenderReplay();
    }

    private void ReplaySpeed_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (ReplaySpeedSelector.SelectedItem is ComboBoxItem item && double.TryParse(item.Tag?.ToString(), out var speed))
            _replaySpeed = speed;
    }

    private void RenderReplay()
    {
        if (!IsLoaded || !ReplayCanvas.IsVisible)
            return;
        ReplayCanvas.Children.Clear();
        if (!TryGetMapProjection(ReplayCanvas, out var projection))
        {
            AddEmptyState(ReplayCanvas, "Load the current wipe map to view route replay.");
            ReplayStateText.Text = "Wipe map unavailable";
            ReplayTimeText.Text = "No map context";
            return;
        }

        DrawWipeGrid(ReplayCanvas, projection);
        if (_points.Count == 0)
        {
            AddEmptyState(ReplayCanvas, "No valid coordinates recorded yet.");
            ReplayStateText.Text = "Waiting for a valid coordinate";
            ReplayTimeText.Text = "No route data";
            return;
        }

        DrawTrack(ReplayCanvas, _points, point => Project(projection, point.X, point.Y), Color.FromArgb(55, 96, 205, 255), 2, 1, dashed: true);
        var visible = _points.Take(Math.Min(_replayIndex + 1, _points.Count)).ToArray();
        DrawTrack(ReplayCanvas, visible, point => Project(projection, point.X, point.Y), Color.FromRgb(96, 205, 255), 3, 1);
        DrawEventMarkers(ReplayCanvas, _points, point => Project(projection, point.X, point.Y), Color.FromRgb(255, 90, 54));
        if (visible.Length > 0)
        {
            var point = Project(projection, visible[^1].X, visible[^1].Y);
            var marker = new Ellipse { Width = 16, Height = 16, Fill = new SolidColorBrush(Color.FromRgb(255, 184, 76)), Stroke = Brushes.White, StrokeThickness = 2 };
            Canvas.SetLeft(marker, point.X - marker.Width / 2);
            Canvas.SetTop(marker, point.Y - marker.Height / 2);
            Panel.SetZIndex(marker, 4);
            ReplayCanvas.Children.Add(marker);
            ReplayStateText.Text = $"{visible[^1].State} · {visible[^1].LocationName ?? visible[^1].LocationType.ToString()}";
            ReplayTimeText.Text = $"{visible[^1].TimestampUtc.ToLocalTime():g} · {visible.Length}/{_points.Count}";
        }
    }

    private void RenderHeatmap()
    {
        if (!IsLoaded || !HeatmapCanvas.IsVisible)
            return;
        HeatmapCanvas.Children.Clear();
        if (!TryGetMapProjection(HeatmapCanvas, out var projection))
        {
            AddEmptyState(HeatmapCanvas, "Load the current wipe map to view movement density.");
            HeatmapText.Text = "Wipe map unavailable";
            return;
        }

        DrawWipeGrid(HeatmapCanvas, projection);
        if (_points.Count == 0)
        {
            AddEmptyState(HeatmapCanvas, "No coordinates recorded yet.");
            HeatmapText.Text = "0 coordinate observations";
            return;
        }

        const double gridSize = 150;
        var columns = Math.Max(1, (int)Math.Ceiling(_worldSize / gridSize));
        var rows = columns;
        var cells = new Dictionary<(int X, int Y), int>();
        foreach (var point in _points)
        {
            var cell = (
                Math.Clamp((int)Math.Floor(point.X / gridSize), 0, columns - 1),
                Math.Clamp((int)Math.Floor((_worldSize - point.Y) / gridSize), 0, rows - 1));
            cells[cell] = cells.TryGetValue(cell, out var count) ? count + 1 : 1;
        }

        var max = cells.Values.DefaultIfEmpty(1).Max();
        var cellTopLeft = Project(projection, 0, _worldSize);
        var cellBottomRight = Project(projection, Math.Min(gridSize, _worldSize), Math.Max(0, _worldSize - gridSize));
        var cellWidth = Math.Max(10, Math.Abs(cellBottomRight.X - cellTopLeft.X) * 1.4);
        var cellHeight = Math.Max(10, Math.Abs(cellBottomRight.Y - cellTopLeft.Y) * 1.4);
        foreach (var (cell, count) in cells)
        {
            var ratio = (double)count / max;
            var cellWest = cell.X * gridSize;
            var cellEast = Math.Min(_worldSize, cellWest + gridSize);
            var cellNorthOffset = cell.Y * gridSize;
            var cellSouthOffset = Math.Min(_worldSize, cellNorthOffset + gridSize);
            var center = Project(
                projection,
                (cellWest + cellEast) / 2,
                _worldSize - (cellNorthOffset + cellSouthOffset) / 2);
            var ellipse = new Ellipse
            {
                Width = cellWidth,
                Height = cellHeight,
                Fill = new SolidColorBrush(HeatColor(ratio)),
                Opacity = 0.3 + ratio * 0.58,
                IsHitTestVisible = false,
            };
            Canvas.SetLeft(ellipse, center.X - ellipse.Width / 2);
            Canvas.SetTop(ellipse, center.Y - ellipse.Height / 2);
            Panel.SetZIndex(ellipse, 2);
            HeatmapCanvas.Children.Add(ellipse);
        }

        DrawTrack(HeatmapCanvas, _points, point => Project(projection, point.X, point.Y), Color.FromArgb(110, 255, 255, 255), 1, 0.65, dashed: true);
        HeatmapText.Text = $"{_points.Count:N0} coordinate observations · hottest 150u grid {max:N0} revisits";
    }

    private void CompareSelector_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (!_refreshing)
            RenderComparison();
    }

    private void RenderComparison()
    {
        if (!IsLoaded || !CompareCanvas.IsVisible)
            return;
        CompareCanvas.Children.Clear();
        DrawGrid(CompareCanvas);
        var a = SelectedPlayerId(CompareASelector) ?? _ownSteamId;
        var b = SelectedPlayerId(CompareBSelector) ?? a;
        var pointsA = ToPoints(_tracker.GetObservations(a));
        var pointsB = ToPoints(_tracker.GetObservations(b));
        var all = pointsA.Concat(pointsB).ToArray();
        if (!TryGetBounds(all, out var bounds))
        {
            AddEmptyState(CompareCanvas, "Two recorded routes are needed for a comparison.");
            CompareSummaryText.Text = "No comparison data";
            return;
        }

        DrawTrack(CompareCanvas, pointsA, point => MapPoint(CompareCanvas, bounds, point.X, point.Y), Color.FromRgb(96, 205, 255), 3, 1);
        DrawTrack(CompareCanvas, pointsB, point => MapPoint(CompareCanvas, bounds, point.X, point.Y), Color.FromRgb(255, 175, 69), 3, 1);
        DrawEventMarkers(CompareCanvas, pointsA, point => MapPoint(CompareCanvas, bounds, point.X, point.Y), Color.FromRgb(96, 205, 255));
        DrawEventMarkers(CompareCanvas, pointsB, point => MapPoint(CompareCanvas, bounds, point.X, point.Y), Color.FromRgb(255, 175, 69));
        CompareSummaryText.Text = $"{ShortName(a)} {pointsA.Count:N0} pts / {Distance(pointsA):N0}u   ·   {ShortName(b)} {pointsB.Count:N0} pts / {Distance(pointsB):N0}u";
    }

    private void CloudRestoreTab_Click(object sender, RoutedEventArgs e)
    {
        TrackerTabs.SelectedItem = CloudTab;
        _ = LoadCloudArchivesAsync();
    }

    private async void RefreshCloud_Click(object sender, RoutedEventArgs e) => await LoadCloudArchivesAsync();

    private async Task LoadCloudArchivesAsync()
    {
        if (!_tracker.Capabilities.CanUseCloudSync)
        {
            CloudStatusText.Text = "Cloud restore is available on premium plans only.";
            return;
        }

        CloudStatusText.Text = "Loading cloud archives…";
        try
        {
            var archives = await _tracker.GetCloudArchivesAsync();
            var items = archives.Select(archive => new CloudArchiveItem(archive)).ToArray();
            CloudArchiveSelector.ItemsSource = items;
            CloudArchiveSelector.SelectedIndex = items.Length > 0 ? 0 : -1;
            CloudStatusText.Text = items.Length == 0 ? "No cloud archives found." : $"{items.Length} archive(s) available.";
        }
        catch (Exception ex)
        {
            CloudStatusText.Text = $"Cloud archive load failed: {ex.Message}";
        }
    }

    private void CloudArchive_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (CloudArchiveSelector.SelectedItem is not CloudArchiveItem item)
        {
            CloudArchiveDetailsText.Text = "Select an archive to inspect it.";
            RestoreCloudButton.IsEnabled = false;
            return;
        }

        CloudArchiveDetailsText.Text = item.Details;
        RestoreCloudButton.IsEnabled = _tracker.Capabilities.CanUseCloudSync;
    }

    private async void RestoreCloud_Click(object sender, RoutedEventArgs e)
    {
        if (CloudArchiveSelector.SelectedItem is not CloudArchiveItem item || !_tracker.Capabilities.CanUseCloudSync)
            return;

        RestoreCloudButton.IsEnabled = false;
        CloudStatusText.Text = "Restoring archive day by day…";
        try
        {
            var result = await _tracker.RestoreCloudArchiveAsync(item.Archive.Id);
            CloudStatusText.Text = result.IsCurrentWipe
                ? $"Restored {result.Observations:N0} observations across {result.Days} day(s). Current views refreshed."
                : $"Restored {result.Observations:N0} observations across {result.Days} day(s). Reconnect to that server/wipe to view the imported history.";
            Refresh();
        }
        catch (Exception ex)
        {
            CloudStatusText.Text = $"Restore failed: {ex.Message}";
        }
        finally
        {
            RestoreCloudButton.IsEnabled = _tracker.Capabilities.CanUseCloudSync;
        }
    }

    private void ExportJson_Click(object sender, RoutedEventArgs e)
    {
        if (!CanExport())
            return;
        var dialog = new SaveFileDialog { Filter = "Tracker JSON (*.json)|*.json", FileName = ExportFileName("json") };
        if (dialog.ShowDialog(this) != true)
            return;

        var steamId = SelectedPlayerId(PlayerSelector) ?? _ownSteamId;
        var document = BuildExportDocument(steamId);
        File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(document, new JsonSerializerOptions { WriteIndented = true }));
        ExportStatusText.Text = $"Exported {_points.Count:N0} observations to {System.IO.Path.GetFileName(dialog.FileName)}.";
    }

    private void ExportCsv_Click(object sender, RoutedEventArgs e)
    {
        if (!CanExport())
            return;
        var dialog = new SaveFileDialog { Filter = "Tracker CSV (*.csv)|*.csv", FileName = ExportFileName("csv") };
        if (dialog.ShowDialog(this) != true)
            return;

        var steamId = SelectedPlayerId(PlayerSelector) ?? _ownSteamId;
        var csv = new StringBuilder("timestamp_utc,steam_id,x,y,state,location_type,location_name,grid,event,session_id\n");
        foreach (var point in ToPoints(_tracker.GetObservations(steamId)))
        {
            csv.Append(string.Join(',',
                Csv(point.TimestampUtc.ToString("O")),
                steamId,
                point.X.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture),
                point.Y.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture),
                Csv(point.State.ToString().ToLowerInvariant()),
                Csv(point.LocationType.ToString().ToLowerInvariant()),
                Csv(point.LocationName),
                Csv(point.Grid),
                Csv(point.Event),
                Csv(point.SessionId)));
            csv.Append('\n');
        }
        File.WriteAllText(dialog.FileName, csv.ToString(), Encoding.UTF8);
        ExportStatusText.Text = $"Exported {_points.Count:N0} observations to {System.IO.Path.GetFileName(dialog.FileName)}.";
    }

    private bool CanExport()
    {
        if (_tracker.Capabilities.CanExport)
            return true;
        ExportStatusText.Text = "Export is available on premium plans only.";
        return false;
    }

    private object BuildExportDocument(ulong steamId)
    {
        var summary = _tracker.GetSummary(steamId);
        return new
        {
            schema_version = 1,
            generated_at = DateTime.UtcNow.ToString("O"),
            player_steam_id = steamId.ToString(),
            player_name = _tracker.GetPlayerName(steamId),
            wipe_key = _tracker.CurrentWipeKey,
            summary = new
            {
                coverage_seconds = (long)summary.Coverage.TotalSeconds,
                unknown_seconds = (long)summary.Unknown.TotalSeconds,
                moving_seconds = (long)summary.Moving.TotalSeconds,
                stationary_seconds = (long)summary.Stationary.TotalSeconds,
                afk_seconds = (long)summary.Afk.TotalSeconds,
                dead_seconds = (long)summary.Dead.TotalSeconds,
                offline_seconds = (long)summary.Offline.TotalSeconds,
                estimated_distance = summary.EstimatedDistance,
                deaths = summary.Deaths,
            },
            observations = ToPoints(_tracker.GetObservations(steamId)),
            segments = _tracker.GetSegments(steamId),
            monument_visits = summary.MonumentVisits,
        };
    }

    private string ExportFileName(string extension)
        => $"player-wipe-{(SelectedPlayerId(PlayerSelector) ?? _ownSteamId)}-{DateTime.Now:yyyyMMdd-HHmm}.{extension}";

    private static string Csv(string? value)
    {
        var text = value ?? string.Empty;
        return text.Contains(',') || text.Contains('"') || text.Contains('\n')
            ? $"\"{text.Replace("\"", "\"\"", StringComparison.Ordinal)}\""
            : text;
    }

    private string ShortName(ulong steamId)
    {
        var name = _tracker.GetPlayerName(steamId);
        return name.Length > 16 ? name[..16] : name;
    }

    private static IReadOnlyList<TrackerPoint> ToPoints(IReadOnlyList<PlayerObservation> observations)
    {
        var points = new List<TrackerPoint>();
        PlayerObservation? previous = null;
        foreach (var observation in observations)
        {
            var state = PlayerWipeTrackerEngine.Classify(observation);
            if (previous is not null && observation.X is not null && observation.Y is not null && previous.X is not null && previous.Y is not null)
            {
                var distance = Math.Sqrt(Math.Pow(observation.X.Value - previous.X.Value, 2) + Math.Pow(observation.Y.Value - previous.Y.Value, 2));
                state = PlayerWipeTrackerEngine.Classify(observation, distance);
            }
            var eventName = previous is not null && !previous.Dead && observation.Dead ? "death" :
                previous is not null && previous.Dead && !observation.Dead ? "respawn" : null;
            if (observation.X is not null && observation.Y is not null)
            {
                points.Add(new TrackerPoint(
                    observation.TimestampUtc,
                    observation.X.Value,
                    observation.Y.Value,
                    state,
                    observation.LocationType,
                    observation.LocationName,
                    observation.Grid,
                    eventName,
                    observation.SessionId));
            }
            previous = observation;
        }
        return points;
    }

    private static double Distance(IReadOnlyList<TrackerPoint> points)
    {
        var distance = 0d;
        for (var i = 1; i < points.Count; i++)
        {
            if (points[i].SessionId != points[i - 1].SessionId)
                continue;
            distance += Math.Sqrt(Math.Pow(points[i].X - points[i - 1].X, 2) + Math.Pow(points[i].Y - points[i - 1].Y, 2));
        }
        return distance;
    }

    private static bool TryGetBounds(IEnumerable<TrackerPoint> points, out MapBounds bounds)
    {
        var list = points.ToArray();
        if (list.Length == 0)
        {
            bounds = default;
            return false;
        }
        var minX = list.Min(point => point.X);
        var maxX = list.Max(point => point.X);
        var minY = list.Min(point => point.Y);
        var maxY = list.Max(point => point.Y);
        var width = Math.Max(1, maxX - minX);
        var height = Math.Max(1, maxY - minY);
        var padX = width * 0.08;
        var padY = height * 0.08;
        bounds = new MapBounds(minX - padX, maxX + padX, minY - padY, maxY + padY);
        return true;
    }

    private static Point Normalize(MapBounds bounds, double x, double y)
        => new((x - bounds.MinX) / Math.Max(1, bounds.MaxX - bounds.MinX),
            (bounds.MaxY - y) / Math.Max(1, bounds.MaxY - bounds.MinY));

    private static Point MapPoint(Canvas canvas, MapBounds bounds, double x, double y)
    {
        var normalized = Normalize(bounds, x, y);
        var width = Math.Max(1, canvas.ActualWidth - 32);
        var height = Math.Max(1, canvas.ActualHeight - 32);
        return new Point(16 + normalized.X * width, 16 + normalized.Y * height);
    }

    private bool TryGetMapProjection(Canvas canvas, out TrackerMapProjection projection)
    {
        var imageWidth = _wipeMapImage?.Width ?? 0;
        var imageHeight = _wipeMapImage?.Height ?? 0;
        var worldRect = _worldRectPixels.Width > 0 && _worldRectPixels.Height > 0
            ? _worldRectPixels
            : new Rect(0, 0, imageWidth, imageHeight);
        projection = new TrackerMapProjection(
            canvas.ActualWidth,
            canvas.ActualHeight,
            imageWidth,
            imageHeight,
            worldRect.X,
            worldRect.Y,
            worldRect.Width,
            worldRect.Height,
            _worldSize);
        return projection.IsValid;
    }

    private static Point Project(TrackerMapProjection projection, double worldX, double worldY)
    {
        var point = projection.Project(worldX, worldY);
        return new Point(point.X, point.Y);
    }

    private void DrawWipeGrid(Canvas canvas, TrackerMapProjection projection)
    {
        const double gridSize = 150;
        var cells = Math.Max(1, (int)Math.Ceiling(_worldSize / gridSize));
        var gridBrush = new SolidColorBrush(Color.FromArgb(100, 225, 240, 245));
        var labelBrush = new SolidColorBrush(Color.FromArgb(210, 240, 248, 250));

        for (var i = 0; i <= cells; i++)
        {
            var world = Math.Min(_worldSize, i * gridSize);
            var verticalTop = Project(projection, world, _worldSize);
            var verticalBottom = Project(projection, world, 0);
            var horizontalLeft = Project(projection, 0, _worldSize - world);
            var horizontalRight = Project(projection, _worldSize, _worldSize - world);
            canvas.Children.Add(new Line { X1 = verticalTop.X, X2 = verticalBottom.X, Y1 = verticalTop.Y, Y2 = verticalBottom.Y, Stroke = gridBrush, StrokeThickness = 0.7, IsHitTestVisible = false });
            canvas.Children.Add(new Line { X1 = horizontalLeft.X, X2 = horizontalRight.X, Y1 = horizontalLeft.Y, Y2 = horizontalRight.Y, Stroke = gridBrush, StrokeThickness = 0.7, IsHitTestVisible = false });
        }

        var firstCell = Project(projection, 0, _worldSize);
        var secondCell = Project(projection, Math.Min(gridSize, _worldSize), Math.Max(0, _worldSize - gridSize));
        var displayedCellSize = Math.Min(Math.Abs(secondCell.X - firstCell.X), Math.Abs(secondCell.Y - firstCell.Y));
        var labelStep = displayedCellSize < 20 ? 3 : displayedCellSize < 30 ? 2 : 1;
        for (var column = 0; column < cells; column += labelStep)
        {
            for (var row = 0; row < cells; row += labelStep)
            {
                var position = Project(
                    projection,
                    Math.Min(_worldSize, column * gridSize + 5),
                    Math.Max(0, _worldSize - row * gridSize - 5));
                var label = new TextBlock
                {
                    Text = $"{ColumnLabel(column)}{row}",
                    Foreground = labelBrush,
                    FontSize = 9,
                    FontWeight = FontWeights.SemiBold,
                    IsHitTestVisible = false,
                };
                Canvas.SetLeft(label, position.X + 2);
                Canvas.SetTop(label, position.Y + 1);
                Panel.SetZIndex(label, 1);
                canvas.Children.Add(label);
            }
        }
    }

    private static string ColumnLabel(int index)
    {
        var label = string.Empty;
        for (index++; index > 0; index /= 26)
        {
            index--;
            label = (char)('A' + index % 26) + label;
        }
        return label;
    }

    private static void DrawGrid(Canvas canvas)
    {
        var width = Math.Max(1, canvas.ActualWidth);
        var height = Math.Max(1, canvas.ActualHeight);
        for (var i = 1; i < 5; i++)
        {
            var x = width * i / 5;
            var y = height * i / 5;
            canvas.Children.Add(new Line { X1 = x, X2 = x, Y1 = 0, Y2 = height, Stroke = new SolidColorBrush(Color.FromArgb(35, 130, 180, 200)), StrokeThickness = 1 });
            canvas.Children.Add(new Line { X1 = 0, X2 = width, Y1 = y, Y2 = y, Stroke = new SolidColorBrush(Color.FromArgb(35, 130, 180, 200)), StrokeThickness = 1 });
        }
    }

    private static void DrawTrack(Canvas canvas, IReadOnlyList<TrackerPoint> points, Func<TrackerPoint, Point> mapPoint, Color color, double thickness, double opacity, bool dashed = false)
    {
        if (points.Count < 2)
            return;
        var batch = new List<Point>();
        TrackerPoint? previous = null;
        foreach (var point in points)
        {
            if (previous is not null && (point.SessionId != previous.SessionId || point.TimestampUtc - previous.TimestampUtc > TimeSpan.FromSeconds(PlayerWipeTrackerEngine.MaxContinuityGapSeconds)))
            {
                AddPolyline(canvas, batch, color, thickness, opacity, dashed);
                batch.Clear();
            }
            batch.Add(mapPoint(point));
            previous = point;
        }
        AddPolyline(canvas, batch, color, thickness, opacity, dashed);
    }

    private static void AddPolyline(Canvas canvas, IReadOnlyList<Point> points, Color color, double thickness, double opacity, bool dashed)
    {
        if (points.Count < 2)
            return;
        var line = new Polyline
        {
            Points = new PointCollection(points),
            Stroke = new SolidColorBrush(color),
            StrokeThickness = thickness,
            Opacity = opacity,
            StrokeLineJoin = PenLineJoin.Round,
            IsHitTestVisible = false,
        };
        if (dashed)
            line.StrokeDashArray = new DoubleCollection { 3, 4 };
        Panel.SetZIndex(line, 3);
        canvas.Children.Add(line);
    }

    private static void DrawEventMarkers(Canvas canvas, IReadOnlyList<TrackerPoint> points, Func<TrackerPoint, Point> mapPoint, Color color)
    {
        foreach (var point in points.Where(point => point.Event == "death"))
        {
            var marker = new Ellipse { Width = 9, Height = 9, Fill = new SolidColorBrush(color), Stroke = Brushes.White, StrokeThickness = 1, IsHitTestVisible = false };
            var position = mapPoint(point);
            Canvas.SetLeft(marker, position.X - marker.Width / 2);
            Canvas.SetTop(marker, position.Y - marker.Height / 2);
            Panel.SetZIndex(marker, 5);
            canvas.Children.Add(marker);
        }
    }

    private static Color HeatColor(double ratio)
    {
        ratio = Math.Clamp(ratio, 0, 1);
        return ratio < 0.5
            ? Blend(Color.FromRgb(22, 63, 106), Color.FromRgb(39, 200, 178), ratio * 2)
            : Blend(Color.FromRgb(39, 200, 178), Color.FromRgb(255, 90, 54), (ratio - 0.5) * 2);
    }

    private static Color Blend(Color from, Color to, double amount)
        => Color.FromRgb(
            (byte)(from.R + (to.R - from.R) * amount),
            (byte)(from.G + (to.G - from.G) * amount),
            (byte)(from.B + (to.B - from.B) * amount));

    private static void AddEmptyState(Canvas canvas, string message)
    {
        var text = new TextBlock { Text = message, Foreground = new SolidColorBrush(Color.FromArgb(180, 220, 230, 235)), FontSize = 14 };
        text.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        Canvas.SetLeft(text, Math.Max(16, (canvas.ActualWidth - text.DesiredSize.Width) / 2));
        Canvas.SetTop(text, Math.Max(16, (canvas.ActualHeight - text.DesiredSize.Height) / 2));
        canvas.Children.Add(text);
    }

    private static string Format(TimeSpan value)
        => value.TotalHours >= 1 ? $"{(int)value.TotalHours}h {value.Minutes}m" : $"{value.Minutes}m {value.Seconds}s";

    private static string FormatBytes(long bytes)
        => bytes >= 1024 * 1024 ? $"{bytes / 1024d / 1024d:N1} MB" : $"{bytes / 1024d:N1} KB";

    private readonly record struct MapBounds(double MinX, double MaxX, double MinY, double MaxY);

    private sealed record PlayerItem(ulong SteamId, string Name)
    {
        public override string ToString() => Name == SteamId.ToString() ? SteamId.ToString() : $"{Name} · {SteamId}";
    }

    private sealed class CloudArchiveItem
    {
        public CloudArchiveItem(CloudArchiveSummary archive) => Archive = archive;
        public CloudArchiveSummary Archive { get; }
        public string Details => $"{Archive.ServerName}\nWipe: {Archive.WipeStartedAtUtc?.ToLocalTime().ToString("g") ?? "unknown"}\nObserved: {Archive.FirstObservedAtUtc?.ToLocalTime().ToString("g") ?? "unknown"} → {Archive.LastObservedAtUtc?.ToLocalTime().ToString("g") ?? "unknown"}\nPlayers: {Archive.PlayerCount?.ToString() ?? "unknown"} · Stored: {FormatBytes(Archive.StoredBytes ?? 0)}";
        public override string ToString() => $"{Archive.ServerName} · {Archive.WipeStartedAtUtc?.ToLocalTime():g} · {Archive.PlayerCount ?? 0} player(s)";
    }
}
