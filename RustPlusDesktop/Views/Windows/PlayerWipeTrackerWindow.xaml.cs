using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using RustPlusDesk.Services.PlayerWipeTracker;

namespace RustPlusDesk.Views.Windows;

public partial class PlayerWipeTrackerWindow : Window
{
    private readonly PlayerWipeTrackerService _tracker;
    private readonly ulong _ownSteamId;

    public PlayerWipeTrackerWindow(PlayerWipeTrackerService tracker, ulong ownSteamId)
    {
        _tracker = tracker ?? throw new ArgumentNullException(nameof(tracker));
        _ownSteamId = ownSteamId;
        InitializeComponent();
        PlayerSelector.Items.Add(new PlayerItem(_ownSteamId, "Self"));
        foreach (var steamId in _tracker.TrackedPlayers.Where(id => id != _ownSteamId))
            PlayerSelector.Items.Add(new PlayerItem(steamId, steamId.ToString()));
        PlayerSelector.SelectedIndex = 0;
        Refresh();
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => Refresh();
    private void PlayerSelector_Changed(object sender, SelectionChangedEventArgs e) => Refresh();

    private void Refresh()
    {
        PremiumBanner.Text = _tracker.Capabilities.CanUseAdvancedViews
            ? "Premium tracker capabilities are active. Cloud backup remains opt-in."
            : "Free tracker: self-only local history. Premium analytics, team tracking, route replay, heatmap, comparison, and export are locked.";
        var steamId = (PlayerSelector.SelectedItem as PlayerItem)?.SteamId ?? _ownSteamId;
        var summary = _tracker.GetSummary(steamId);
        CoverageText.Text = Format(summary.Coverage);
        UnknownText.Text = Format(summary.Unknown);
        DistanceText.Text = $"{summary.EstimatedDistance:N0} units";
        DeathsText.Text = summary.Deaths.ToString();
        StorageText.Text = FormatBytes(_tracker.StorageBytes);
        TimelineList.ItemsSource = _tracker.GetSegments(steamId)
            .OrderByDescending(s => s.StartUtc)
            .Select(s => new
            {
                Start = s.StartUtc.ToLocalTime().ToString("g"),
                End = s.EndUtc.ToLocalTime().ToString("g"),
                State = s.State.ToString(),
                Location = s.LocationName ?? s.LocationType.ToString(),
            }).ToArray();
        VisitsList.ItemsSource = summary.MonumentVisits
            .OrderByDescending(v => v.StartUtc)
            .Select(v => new { v.Name, Duration = Format(v.EstimatedDuration) })
            .ToArray();
    }

    private static string Format(TimeSpan value)
        => value.TotalHours >= 1 ? $"{(int)value.TotalHours}h {value.Minutes}m" : $"{value.Minutes}m {value.Seconds}s";

    private static string FormatBytes(long bytes)
        => bytes >= 1024 * 1024 ? $"{bytes / 1024d / 1024d:N1} MB" : $"{bytes / 1024d:N1} KB";

    private sealed record PlayerItem(ulong SteamId, string Name)
    {
        public override string ToString() => Name == "Self" ? $"{Name} ({SteamId})" : Name;
    }
}
