using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Diagnostics;

namespace RustPlusDesk.Avalonia.Features.Diagnostics;

/// <summary>
/// Reference TrafficMonitorWindow: the live speeds, the totals, the per-request log and the category breakdown.
/// The monitor itself is the app's own traffic - nothing about it leaves the machine.
/// </summary>
public sealed partial class TrafficMonitorViewModel(NetworkTrafficMonitor monitor) : ObservableObject
{
    public ObservableCollection<TrafficEntry> Entries { get; } = [];
    public ObservableCollection<CategoryBandwidthStat> Categories { get; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Summary))]
    private string downloadSpeed = "0 bps";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Summary))]
    private string uploadSpeed = "0 bps";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Summary))]
    private string total = "0 B";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Summary))]
    private int requestCount;

    /// <summary>
    /// The one-line readout the title bar shows (reference BandwidthSummaryBadge): the live down and up rates, or "Off"
    /// while the monitor is switched off.
    /// </summary>
    public string Summary => monitor.IsEnabled ? $"{DownloadSpeed} ↓ {UploadSpeed} ↑" : "Off";

    /// <summary>Reference IsRecording: the log can be paused without losing what is already there.</summary>
    public bool IsRecording
    {
        get => monitor.IsRecording;
        set
        {
            if (monitor.IsRecording == value) return;
            monitor.IsRecording = value;
            OnPropertyChanged();
            Status = value ? "Recording." : "Paused: new requests are not recorded.";
        }
    }

    /// <summary>Reference Clear: the counters and the log start over.</summary>
    [RelayCommand]
    public void Clear()
    {
        monitor.Clear();
        Entries.Clear();
        Categories.Clear();
        TotalsChanged();
    }

    /// <summary>Reference the export: the log is written as a comma separated file the user can open anywhere.</summary>
    [RelayCommand]
    public async Task ExportAsync()
    {
        var rows = monitor.Entries;
        if (rows.Count == 0) return;
        string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            $"rustplus-desktop-traffic-{DateTime.Now:yyyyMMdd-HHmmss}.csv");
        try
        {
            var lines = new List<string> { "Time,Category,Direction,Endpoint,BytesIn,BytesOut,DurationMs,Status,Details" };
            foreach (var row in rows)
                lines.Add(string.Join(',', [
                    row.TimestampFormatted,
                    Escape(row.Category),
                    Escape(row.DirectionText),
                    Escape(row.Endpoint),
                    row.BytesIn.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    row.BytesOut.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    row.DurationMs.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Escape(row.Status),
                    Escape(row.Details)]));
            await File.WriteAllLinesAsync(path, lines);
            Status = $"Written to {path}";
        }
        catch (Exception ex) { Status = $"The log could not be written: {ex.Message}"; }
    }

    [ObservableProperty] private string status = "";

    private static string Escape(string text) => "\"" + (text ?? "").Replace("\"", "\"\"") + "\"";

    /// <summary>Reads the monitor and fills the lists.</summary>
    public void Refresh()
    {
        Entries.Clear();
        foreach (var entry in monitor.Entries) Entries.Add(entry);
        Categories.Clear();
        foreach (var category in monitor.CategoryBreakdown) Categories.Add(category);
        DownloadSpeed = monitor.FormattedDownloadSpeedBits;
        UploadSpeed = monitor.FormattedUploadSpeedBits;
        Total = monitor.FormattedTotalBytes;
        RequestCount = (int)monitor.TotalRequestsCount;
    }

    private void TotalsChanged()
    {
        DownloadSpeed = monitor.FormattedDownloadSpeedBits;
        UploadSpeed = monitor.FormattedUploadSpeedBits;
        Total = monitor.FormattedTotalBytes;
        RequestCount = (int)monitor.TotalRequestsCount;
    }
}
