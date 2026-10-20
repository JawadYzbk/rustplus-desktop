using System.Collections.Concurrent;
using System.Globalization;

namespace RustPlusDesk.Core.Diagnostics;

/// <summary>Which way the bytes went (reference TrafficDirection).</summary>
public enum TrafficDirection { Inbound, Outbound, Both }

/// <summary>One recorded request (reference TrafficEntry).</summary>
public sealed record TrafficEntry(long Id, DateTime Timestamp, string Category, TrafficDirection Direction,
    string Endpoint, long BytesIn, long BytesOut, long DurationMs, string Status, string Details)
{
    public long TotalBytes => BytesIn + BytesOut;

    public string FormattedSize => NetworkTrafficMonitor.FormatBytes(TotalBytes);

    public string TimestampFormatted => Timestamp.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);

    public string DirectionText => Direction switch
    {
        TrafficDirection.Inbound => "↓ IN",
        TrafficDirection.Outbound => "↑ OUT",
        _ => "⇅"
    };
}

/// <summary>How much one category used (reference CategoryBandwidthStat).</summary>
public sealed record CategoryBandwidthStat(string Category, long TotalBytesIn, long TotalBytesOut, int RequestCount,
    double PercentageOfTotal)
{
    public long TotalBytes => TotalBytesIn + TotalBytesOut;
    public string FormattedTotalBytes => NetworkTrafficMonitor.FormatBytes(TotalBytes);
    public string FormattedBytesIn => NetworkTrafficMonitor.FormatBytes(TotalBytesIn);
    public string FormattedBytesOut => NetworkTrafficMonitor.FormatBytes(TotalBytesOut);
    public string FormattedPercentage => $"{PercentageOfTotal:F1}%";
}

/// <summary>
/// Reference NetworkTrafficMonitor: the app's own network traffic, counted and shown.
///
/// Everything is local: the app measures its own requests, so nothing about the user's traffic leaves the machine.
/// The log is bounded, the speeds are a rolling window over the last second and a half, and the category breakdown
/// says where the bytes went.
/// </summary>
public sealed class NetworkTrafficMonitor
{
    public const int MaxLogEntries = 2500;

    private readonly object gate = new();
    private readonly ConcurrentQueue<(DateTime Time, long InBytes, long OutBytes)> speedHistory = new();
    private readonly Dictionary<string, CategoryBandwidthStat> categoryStats = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<TrafficEntry> entries = [];
    private long nextId = 1;
    private long totalInboundBytes;
    private long totalOutboundBytes;
    private long totalRequestsCount;

    /// <summary>Raised for every recorded request, so the view can add it to the log.</summary>
    public event Action<TrafficEntry>? EntryRecorded;

    /// <summary>Raised when the totals change, so the readout follows them.</summary>
    public event Action? TotalsChanged;

    public bool IsEnabled { get; set; } = true;

    public bool IsRecording { get; set; } = true;

    public long TotalInboundBytes { get { lock (gate) return totalInboundBytes; } }
    public long TotalOutboundBytes { get { lock (gate) return totalOutboundBytes; } }
    public long TotalBytes { get { lock (gate) return totalInboundBytes + totalOutboundBytes; } }
    public long TotalRequestsCount { get { lock (gate) return totalRequestsCount; } }

    public string FormattedTotalInbound => FormatBytes(TotalInboundBytes);
    public string FormattedTotalOutbound => FormatBytes(TotalOutboundBytes);
    public string FormattedTotalBytes => FormatBytes(TotalBytes);

    private double currentDownloadSpeedBps;
    private double currentUploadSpeedBps;

    /// <summary>The rolling download speed, in bytes per second.</summary>
    public double CurrentDownloadSpeedBps { get { lock (gate) return currentDownloadSpeedBps; } }

    /// <summary>The rolling upload speed, in bytes per second.</summary>
    public double CurrentUploadSpeedBps { get { lock (gate) return currentUploadSpeedBps; } }

    public string FormattedDownloadSpeed => FormatBytes((long)CurrentDownloadSpeedBps) + "/s";
    public string FormattedUploadSpeed => FormatBytes((long)CurrentUploadSpeedBps) + "/s";
    public string FormattedDownloadSpeedBits => FormatBits((long)CurrentDownloadSpeedBps);
    public string FormattedUploadSpeedBits => FormatBits((long)CurrentUploadSpeedBps);

    /// <summary>The recorded requests, newest last.</summary>
    public IReadOnlyList<TrafficEntry> Entries { get { lock (gate) return entries.ToArray(); } }

    /// <summary>What each category used, biggest first.</summary>
    public IReadOnlyList<CategoryBandwidthStat> CategoryBreakdown
    {
        get
        {
            lock (gate)
            {
                long total = totalInboundBytes + totalOutboundBytes;
                if (total <= 0) total = 1;
                return categoryStats.Values
                    .Select(stat => stat with { PercentageOfTotal = (double)stat.TotalBytes / total * 100.0 })
                    .OrderByDescending(stat => stat.TotalBytes)
                    .ToArray();
            }
        }
    }

    /// <summary>Records one request (reference NetworkTrafficMonitor.Record).</summary>
    public void Record(string category, TrafficDirection direction, string endpoint, long bytesIn, long bytesOut,
        long durationMs = 0, string status = "OK", string details = "")
    {
        if (!IsEnabled || !IsRecording) return;

        string effectiveCategory = string.IsNullOrWhiteSpace(category) ? "General" : category;
        var entry = new TrafficEntry(
            Interlocked.Increment(ref nextId),
            DateTime.Now,
            effectiveCategory,
            direction,
            endpoint,
            Math.Max(0, bytesIn),
            Math.Max(0, bytesOut),
            durationMs,
            status,
            details);

        speedHistory.Enqueue((DateTime.UtcNow, entry.BytesIn, entry.BytesOut));
        lock (gate)
        {
            totalInboundBytes += entry.BytesIn;
            totalOutboundBytes += entry.BytesOut;
            totalRequestsCount++;

            var stat = categoryStats.GetValueOrDefault(effectiveCategory);
            if (stat is null)
            {
                stat = new CategoryBandwidthStat(effectiveCategory, 0, 0, 0, 0);
                categoryStats[effectiveCategory] = stat;
            }
            categoryStats[effectiveCategory] = stat with
            {
                TotalBytesIn = stat.TotalBytesIn + entry.BytesIn,
                TotalBytesOut = stat.TotalBytesOut + entry.BytesOut,
                RequestCount = stat.RequestCount + 1
            };

            entries.Add(entry);
            if (entries.Count > MaxLogEntries) entries.RemoveAt(0);
        }

        TotalsChanged?.Invoke();
        EntryRecorded?.Invoke(entry);
    }

    public void RecordInbound(string category, string endpoint, long bytesIn, long durationMs = 0, string status = "OK",
        string details = "") => Record(category, TrafficDirection.Inbound, endpoint, bytesIn, 0, durationMs, status, details);

    public void RecordOutbound(string category, string endpoint, long bytesOut, long durationMs = 0, string status = "Sent",
        string details = "") => Record(category, TrafficDirection.Outbound, endpoint, 0, bytesOut, durationMs, status, details);

    /// <summary>Reference Clear: the counters and the log start over.</summary>
    public void Clear()
    {
        lock (gate)
        {
            categoryStats.Clear();
            entries.Clear();
            totalInboundBytes = 0;
            totalOutboundBytes = 0;
            totalRequestsCount = 0;
            currentDownloadSpeedBps = 0;
            currentUploadSpeedBps = 0;
        }
        while (speedHistory.TryDequeue(out _)) { }
        TotalsChanged?.Invoke();
    }

    /// <summary>
    /// Reference SpeedTimer_Tick: the speed is what came in over the last second and a half, so a burst that already
    /// ended does not keep showing.
    /// </summary>
    public void TickSpeed()
    {
        if (!IsEnabled)
        {
            lock (gate) { currentDownloadSpeedBps = 0; currentUploadSpeedBps = 0; }
            TotalsChanged?.Invoke();
            return;
        }

        DateTime windowStart = DateTime.UtcNow.AddSeconds(-1.5);
        while (speedHistory.TryPeek(out var item) && item.Time < windowStart) speedHistory.TryDequeue(out _);

        long totalIn = 0, totalOut = 0;
        foreach (var sample in speedHistory) { totalIn += sample.InBytes; totalOut += sample.OutBytes; }

        lock (gate)
        {
            currentDownloadSpeedBps = totalIn / 1.5;
            currentUploadSpeedBps = totalOut / 1.5;
        }
        TotalsChanged?.Invoke();
    }

    /// <summary>Reference FormatBytes.</summary>
    public static string FormatBytes(long bytes)
    {
        if (bytes <= 0) return "0 B";
        string[] units = ["B", "KB", "MB", "GB", "TB"];
        double value = bytes;
        int unitIndex = 0;
        while (value >= 1024 && unitIndex < units.Length - 1)
        {
            value /= 1024;
            unitIndex++;
        }
        if (unitIndex == 0) return $"{bytes} B";
        return $"{value:0.#} {units[unitIndex]}";
    }

    /// <summary>Reference FormatBits.</summary>
    public static string FormatBits(long bytes)
    {
        if (bytes <= 0) return "0 bps";
        double bits = bytes * 8.0;
        string[] units = ["bps", "Kbps", "Mbps", "Gbps", "Tbps"];
        int unitIndex = 0;
        while (bits >= 1000.0 && unitIndex < units.Length - 1)
        {
            bits /= 1000.0;
            unitIndex++;
        }
        if (unitIndex == 0) return $"{bits:0} bps";
        return $"{bits:0.##} {units[unitIndex]}";
    }
}
