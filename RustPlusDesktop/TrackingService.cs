using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace RustPlusDesk.Services;

public class TrackedPlayer
{
    public string BMId { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string LastServerName { get; set; } = string.Empty;
    public List<PlayerSession> Sessions { get; set; } = new();
}

public class TrackingSettings
{
    public string LastHost { get; set; } = string.Empty;
    public int LastPort { get; set; } = 0;
    public string LastServerName { get; set; } = string.Empty;
    public bool BackgroundTrackingEnabled { get; set; } = false;
    public bool CloseToTrayEnabled { get; set; } = false;
    public bool StartMinimizedEnabled { get; set; } = false;
    public bool AutoConnectEnabled { get; set; } = false;
    public bool AutoStartEnabled { get; set; } = false;
    public string SteamId64 { get; set; } = string.Empty;
}

public class PlayerSession
{
    public DateTime ConnectTime { get; set; }
    public DateTime? DisconnectTime { get; set; }
}

public class OnlinePlayerBM
{
    public string Name { get; set; } = string.Empty;
    public string BMId { get; set; } = string.Empty;
    public DateTime SessionStartTimeUtc { get; set; }
    public TimeSpan Duration { get; set; }
    public bool IsTracked { get; set; }
    public string PlayTimeStr => $"{(int)Duration.TotalHours:D2}:{Duration.Minutes:D2}";
}

public static class TrackingService
{
    private static readonly HttpClient _http = new();
    private static readonly string _dbPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
        "RustPlusDesk", "tracked_players.json");
    private static readonly string _settingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "tracking_settings.json");
    
    private static Dictionary<string, TrackedPlayer> _trackedPlayers = new();
    private static TrackingSettings _settings = new();
    private static Timer? _trackingTimer;
    private static string? _lastServerHost;
    private static int _lastServerPort;
    private static string? _lastServerName;

    public static event Action? OnOnlinePlayersUpdated;
    public static event Action<string>? OnServerInfoUpdated;
    public static string StatusMessage { get; private set; } = "";
    public static List<OnlinePlayerBM> LastOnlinePlayers { get; private set; } = new();
    public static DateTime? LastPullTime { get; private set; }
    public static bool IsTracking => _trackingTimer != null;

    static TrackingService()
    {
        _http.DefaultRequestHeaders.Add("User-Agent", "RustPlusDesk/1.0");
        LoadDB();
    }

    private static void LoadDB()
    {
        try
        {
            if (File.Exists(_dbPath))
            {
                var json = File.ReadAllText(_dbPath);
                var list = JsonSerializer.Deserialize<List<TrackedPlayer>>(json);
                if (list != null) _trackedPlayers = list.ToDictionary(p => p.BMId);
            }
            if (File.Exists(_settingsPath))
            {
                var json = File.ReadAllText(_settingsPath);
                _settings = JsonSerializer.Deserialize<TrackingSettings>(json) ?? new();
            }
        }
        catch { }
    }

    public static void SaveDB()
    {
        try
        {
            var dir = Path.GetDirectoryName(_dbPath);
            if (dir != null && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

            var jsonP = JsonSerializer.Serialize(_trackedPlayers.Values.ToList(), new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_dbPath, jsonP);

            var jsonS = JsonSerializer.Serialize(_settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_settingsPath, jsonS);
        }
        catch { }
    }

    private static void Log(string message)
    {
        try
        {
            var logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "RustPlusDesk", "tracking_log.txt");
            var dir = Path.GetDirectoryName(logPath);
            if (dir != null && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
        }
        catch { }
    }

    public static void TrackPlayer(string bmId, string name, string serverName, PlayerSession? initialSession = null)
    {
        TrackedPlayer? player;
        if (!_trackedPlayers.TryGetValue(bmId, out player))
        {
            player = new TrackedPlayer { BMId = bmId, Name = name, LastServerName = serverName };
            _trackedPlayers[bmId] = player;
        }
        else
        {
            player.LastServerName = serverName;
            // Update name if we got a real one
            if (name != "Unknown Player") player.Name = name;
        }

        if (initialSession != null)
        {
            // Only add if we don't have overlapping sessions already
            if (!player.Sessions.Any(s => s.ConnectTime == initialSession.ConnectTime))
            {
                player.Sessions.Add(initialSession);
                player.Sessions = player.Sessions.OrderBy(s => s.ConnectTime).ToList();
            }
        }

        SaveDB();

        // Auto-start tracking if we have a server but no timer yet
        if (_trackingTimer == null && !string.IsNullOrEmpty(_settings.LastHost))
        {
            StartPolling(_settings.LastHost, _settings.LastPort, _settings.LastServerName);
        }
        OnOnlinePlayersUpdated?.Invoke();
    }
    
    public static void UntrackPlayer(string bmId)
    {
        if (_trackedPlayers.Remove(bmId))
        {
            SaveDB();
            if (_trackedPlayers.Count == 0)
            {
                StopPolling();
            }
            OnOnlinePlayersUpdated?.Invoke();
        }
    }
    
    public static string? CurrentServerBMId => _foundServerId;

    public static void RenameTrackedPlayer(string bmId, string newName)
    {
        if (_trackedPlayers.TryGetValue(bmId, out var player))
        {
            player.Name = newName;
            SaveDB();
            OnOnlinePlayersUpdated?.Invoke();
        }
    }
    public static List<TrackedPlayer> GetTrackedPlayers() => _trackedPlayers.Values.ToList();
    public static bool IsTracked(string bmId) => _trackedPlayers.ContainsKey(bmId);

    public static bool IsBackgroundTrackingEnabled
    {
        get => _settings.BackgroundTrackingEnabled;
        set { _settings.BackgroundTrackingEnabled = value; SaveDB(); }
    }

    public static bool CloseToTrayEnabled
    {
        get => _settings.CloseToTrayEnabled;
        set { _settings.CloseToTrayEnabled = value; SaveDB(); }
    }

    public static bool StartMinimizedEnabled
    {
        get => _settings.StartMinimizedEnabled;
        set { _settings.StartMinimizedEnabled = value; SaveDB(); }
    }

    public static bool AutoConnectEnabled
    {
        get => _settings.AutoConnectEnabled;
        set { _settings.AutoConnectEnabled = value; SaveDB(); }
    }

    public static bool AutoStartEnabled
    {
        get => _settings.AutoStartEnabled;
        set 
        { 
            if (_settings.AutoStartEnabled == value) return;
            _settings.AutoStartEnabled = value; 
            SetAutoStart(value);
            SaveDB(); 
        }
    }

    public static string SteamId64
    {
        get => _settings.SteamId64;
        set { _settings.SteamId64 = value; SaveDB(); }
    }

    private static void SetAutoStart(bool enabled)
    {
        try
        {
            const string runKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(runKey, true);
            if (key == null) return;

            string appName = "RustPlusDesk";
            if (enabled)
            {
                string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule!.FileName!;
                key.SetValue(appName, $"\"{exePath}\" --background");
            }
            else
            {
                key.DeleteValue(appName, false);
            }
        }
        catch { }
    }

    public static (string host, int port, string name) LastServer => (_settings.LastHost, _settings.LastPort, _settings.LastServerName);

    public static async Task<string> FetchPlayerNameAsync(string bmId)
    {
        try
        {
            // 1. Try direct player endpoint
            var url = $"https://api.battlemetrics.com/players/{bmId}";
            var response = await _http.GetAsync(url);
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                return doc.RootElement.GetProperty("data").GetProperty("attributes").GetProperty("name").GetString() ?? "Unknown Player";
            }
            
            Log($"[API] FetchPlayerName direct failed for {bmId}: {response.StatusCode}. Trying session fallback...");
            
            // 2. Fallback: Try most recent session to get name
            var sUrl = $"https://api.battlemetrics.com/sessions?filter[players]={bmId}&page[size]=1";
            var sResponse = await _http.GetAsync(sUrl);
            if (sResponse.IsSuccessStatusCode)
            {
                var sJson = await sResponse.Content.ReadAsStringAsync();
                using var sDoc = JsonDocument.Parse(sJson);
                if (sDoc.RootElement.TryGetProperty("data", out var sData) && sData.ValueKind == JsonValueKind.Array && sData.GetArrayLength() > 0)
                {
                    return sData[0].GetProperty("attributes").GetProperty("name").GetString() ?? "Unknown Player";
                }
            }
            
            return "Unknown Player";
        }
        catch (Exception ex) 
        { 
            Log($"[API] Error fetching name for {bmId}: {ex.Message}");
            return "Unknown Player"; 
        }
    }

    public static async Task<DateTime?> FetchPlayerLastSeenAsync(string bmId)
    {
        if (string.IsNullOrEmpty(_foundServerId)) return null;
        try
        {
            var url = $"https://api.battlemetrics.com/players/{bmId}/servers/{_foundServerId}";
            var response = await _http.GetAsync(url);
            if (!response.IsSuccessStatusCode) return null;

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            
            if (doc.RootElement.TryGetProperty("data", out var data) && 
                data.TryGetProperty("attributes", out var attr))
            {
                if (attr.TryGetProperty("lastSeen", out var stopProp) && stopProp.ValueKind == JsonValueKind.String)
                {
                    if (DateTimeOffset.TryParse(stopProp.GetString(), out var stop))
                    {
                        return stop.UtcDateTime;
                    }
                }
            }
        }
        catch { }
        return null;
    }

    public static void LoadDemoData()
    {
        _trackedPlayers.Clear();
        var now = DateTime.UtcNow;

        // 1. The Night Owl (Plays 00:00 - 06:00)
        var owl = new TrackedPlayer { BMId = "demo_1", Name = "NightOwl_X" };
        for (int d = 0; d < 14; d++) {
            var date = now.Date.AddDays(-d).AddHours(1); // 01:00
            owl.Sessions.Add(new PlayerSession { ConnectTime = date, DisconnectTime = date.AddHours(4) });
        }
        _trackedPlayers[owl.BMId] = owl;

        // 2. The Grinder (Huge playtime, active 12:00 - 02:00)
        var grinder = new TrackedPlayer { BMId = "demo_2", Name = "IndustrialPvP" };
        for (int d = 0; d < 7; d++) {
            var date = now.Date.AddDays(-d).AddHours(12); // Noon
            grinder.Sessions.Add(new PlayerSession { ConnectTime = date, DisconnectTime = date.AddHours(14) }); // Until 02:00
        }
        _trackedPlayers[grinder.BMId] = grinder;

        // 3. The Weekend Warrior (Only Sat/Sun)
        var weekend = new TrackedPlayer { BMId = "demo_3", Name = "CasualFriday" };
        for (int d = 0; d < 30; d++) {
            var date = now.Date.AddDays(-d);
            if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday) {
                weekend.Sessions.Add(new PlayerSession { ConnectTime = date.AddHours(10), DisconnectTime = date.AddHours(18) });
            }
        }
        _trackedPlayers[weekend.BMId] = weekend;

        SaveDB();
        OnOnlinePlayersUpdated?.Invoke();
    }

    public static async Task<PlayerSession?> FetchPlayerLastSessionAsync(string bmId)
    {
        if (string.IsNullOrEmpty(_foundServerId)) return null;
        try
        {
            // Try to fetch the most recent session for this player on this server
            var url = $"https://api.battlemetrics.com/sessions?filter[players]={bmId}&filter[servers]={_foundServerId}&page[size]=1";
            var response = await _http.GetAsync(url);
            if (!response.IsSuccessStatusCode) return null;

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            
            if (doc.RootElement.TryGetProperty("data", out var data) && data.ValueKind == JsonValueKind.Array && data.GetArrayLength() > 0)
            {
                var sessionObj = data[0];
                var attr = sessionObj.GetProperty("attributes");
                
                DateTime? start = null;
                DateTime? stop = null;

                if (attr.TryGetProperty("start", out var sProp) && sProp.ValueKind == JsonValueKind.String)
                {
                    if (DateTimeOffset.TryParse(sProp.GetString(), out var s)) start = s.UtcDateTime;
                }
                if (attr.TryGetProperty("stop", out var eProp) && eProp.ValueKind == JsonValueKind.String)
                {
                    if (DateTimeOffset.TryParse(eProp.GetString(), out var e)) stop = e.UtcDateTime;
                }

                if (start.HasValue)
                {
                    return new PlayerSession { ConnectTime = start.Value, DisconnectTime = stop };
                }
            }
        }
        catch { }
        return null;
    }
    public static string GetAnalysisReport(string? targetBmId = null)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'>");
        sb.AppendLine("<style>");
        // Root styles
        sb.AppendLine("html { background: #11100e; }");
        sb.AppendLine("body { background: #11100e; color: #f5efe7; font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif; margin: 0; padding: 26px; line-height: 1.45; }");
        sb.AppendLine(".report-shell { max-width: 1120px; margin: 0 auto; }");
        sb.AppendLine(".report-kicker { color: #d66a38; font-size: 11px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; margin-bottom: 6px; }");
        sb.AppendLine(".report-subtitle { color: #b8aaa0; font-size: 13px; margin: -18px 0 24px 0; }");
        sb.AppendLine(".player-card { background: #161310; border: 1px solid #3a2e26; border-radius: 8px; padding: 22px; margin-bottom: 22px; box-shadow: 0 10px 28px rgba(0,0,0,0.24); }");
        sb.AppendLine("h1 { color: #f5efe7; font-size: 28px; font-weight: 700; margin: 0 0 24px 0; letter-spacing: 0; }");

        // Theme variables (to be overridden per card)
        sb.AppendLine(".theme-online { --theme-accent: #62d38b; --theme-accent-soft: rgba(98, 211, 139, 0.10); --theme-accent-border: rgba(98, 211, 139, 0.30); --cell-lv1: #2c3827; --cell-lv2: #476833; --cell-lv3: #76a747; --cell-lv4: #a7d66d; }");
        sb.AppendLine(".theme-offline { --theme-accent: #b8aaa0; --theme-accent-soft: rgba(214, 106, 56, 0.08); --theme-accent-border: rgba(214, 106, 56, 0.22); --cell-lv1: #211c17; --cell-lv2: #3a2e26; --cell-lv3: #773e25; --cell-lv4: #d66a38; }");

        sb.AppendLine("h2 { color: #f5efe7; margin: 0; font-size: 22px; font-weight: 700; }");
        sb.AppendLine(".player-heading { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; border-bottom: 1px solid #3a2e26; padding-bottom: 14px; margin-bottom: 16px; }");
        sb.AppendLine(".player-meta { color: #786a60; font-size: 11px; margin-top: 3px; }");
        sb.AppendLine(".server-title { color: #d66a38; font-size: 12px; font-weight: 700; margin: 28px 0 10px 0; text-transform: uppercase; letter-spacing: .06em; border-bottom: 1px solid #3a2e26; padding-bottom: 8px; }");
        sb.AppendLine(".stat-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-bottom: 20px; }");
        sb.AppendLine(".stat-item { background: #211c17; padding: 12px; border-radius: 8px; border: 1px solid #3a2e26; }");
        sb.AppendLine(".stat-label { font-size: 10px; color: #b8aaa0; text-transform: uppercase; font-weight: 700; }");
        sb.AppendLine(".stat-value { font-size: 16px; color: #f5efe7; font-weight: 700; margin-top: 4px; }");
        
        sb.AppendLine(".badge { display: inline-flex; align-items: center; min-height: 22px; padding: 0 10px; border-radius: 999px; font-size: 11px; font-weight: 700; text-transform: uppercase; }");
        sb.AppendLine(".badge-online { background: rgba(98, 211, 139, 0.10); color: #62d38b; border: 1px solid rgba(98, 211, 139, 0.35); }");
        sb.AppendLine(".badge-offline { background: rgba(184, 170, 160, 0.08); color: #b8aaa0; border: 1px solid rgba(184, 170, 160, 0.18); }");
        
        sb.AppendLine(".section-title { font-size: 12px; font-weight: 700; color: #b8aaa0; margin: 24px 0 10px 0; display: flex; align-items: center; text-transform: uppercase; }");
        sb.AppendLine(".section-title::after { content: ''; flex: 1; height: 1px; background: #3a2e26; margin-left: 10px; }");

        // Activity grid
        sb.AppendLine(".grid-container { display: grid; grid-template-columns: repeat(12, 1fr); gap: 10px; margin-top: 10px; }");
        sb.AppendLine(".grid-week { display: grid; grid-template-rows: repeat(7, 10px); gap: 2px; }");
        sb.AppendLine(".grid-cell { width: 10px; height: 10px; border-radius: 2px; background: #211c17; border: 1px solid rgba(255,255,255,0.03); }");
        sb.AppendLine(".grid-cell.lv1 { background: var(--cell-lv1); }");
        sb.AppendLine(".grid-cell.lv2 { background: var(--cell-lv2); }");
        sb.AppendLine(".grid-cell.lv3 { background: var(--cell-lv3); }");
        sb.AppendLine(".grid-cell.lv4 { background: var(--cell-lv4); }");

        // Hourly heat
        sb.AppendLine(".hourly-wrap { background: #211c17; padding: 15px; border-radius: 8px; border: 1px solid #3a2e26; }");
        sb.AppendLine(".hourly-container { display: flex; height: 60px; gap: 2px; align-items: flex-end; }");
        sb.AppendLine(".hour-bar { flex: 1; background: #3a2e26; border-radius: 3px 3px 0 0; position: relative; }");
        sb.AppendLine(".hour-bar.active { background: var(--theme-accent); }");
        sb.AppendLine(".hour-labels { display: flex; justify-content: space-between; margin-top: 8px; font-size: 10px; color: #786a60; font-family: monospace; }");
        
        sb.AppendLine(".insight-box { background: var(--theme-accent-soft); border: 1px solid var(--theme-accent-border); padding: 16px; margin-top: 20px; border-radius: 8px; }");
        sb.AppendLine(".insight-item { margin: 8px 0; font-size: 14px; display: flex; align-items: center; color: #f5efe7; }");
        sb.AppendLine(".insight-label { color: #b8aaa0; font-weight: 700; min-width: 128px; }");
        sb.AppendLine(".warning { background: rgba(214, 106, 56, 0.10); border: 1px solid rgba(214, 106, 56, 0.24); color: #d66a38; padding: 10px; border-radius: 8px; font-size: 12px; margin-top: 15px; }");
        sb.AppendLine(".empty-state { background: #161310; border: 1px solid #3a2e26; border-radius: 8px; padding: 18px; color: #b8aaa0; }");
        sb.AppendLine("@media (max-width: 760px) { body { padding: 16px; } .stat-grid { grid-template-columns: 1fr; } .player-heading { display: block; } .badge { margin-top: 10px; } }");
        sb.AppendLine("</style></head><body><div class='report-shell'>");

        sb.AppendLine("<div class='report-kicker'>RustPlusDesk tracker</div>");
        sb.AppendLine("<h1>Activity Intelligence Report</h1>");
        sb.AppendLine("<div class='report-subtitle'>Session history, intensity, and likely activity windows for tracked BattleMetrics targets.</div>");
        
        var playersToReport = targetBmId == null 
            ? _trackedPlayers.Values.ToList() 
            : _trackedPlayers.Values.Where(p => p.BMId == targetBmId).ToList();

        if (!playersToReport.Any())
        {
            sb.AppendLine("<div class='empty-state'>No players in tracking database. Start by tracking players from the server list.</div>");
        }

        var groupedPlayers = playersToReport.GroupBy(p => string.IsNullOrEmpty(p.LastServerName) ? "Global / Legacy" : p.LastServerName);

        foreach(var group in groupedPlayers)
        {
            sb.AppendLine($"<div class='server-title'>{WebUtility.HtmlEncode(group.Key)}</div>");
            
            foreach(var p in group)
            {
                var totalTime = TimeSpan.Zero;
                var past7Days = TimeSpan.Zero;
                var now = DateTime.UtcNow;
                
                int[] hourActivity = new int[24];
                Dictionary<DateTime, int> dailyActivity = new Dictionary<DateTime, int>();

                foreach (var session in p.Sessions)
                {
                    var end = session.DisconnectTime ?? now;
                    var dur = end - session.ConnectTime;
                    totalTime += dur;
                    if (session.ConnectTime > now.AddDays(-7)) past7Days += dur;

                    var date = session.ConnectTime.Date;
                    if (!dailyActivity.ContainsKey(date)) dailyActivity[date] = 0;
                    dailyActivity[date] += (int)dur.TotalMinutes;

                    var iter = session.ConnectTime;
                    while (iter < end)
                    {
                        hourActivity[iter.ToLocalTime().Hour]++;
                        iter = iter.AddHours(1);
                    }
                }

                double avgSessionMins = p.Sessions.Any() ? totalTime.TotalMinutes / p.Sessions.Count : 0;
                var isOnline = p.Sessions.Any() && !p.Sessions.Last().DisconnectTime.HasValue;
                var themeClass = isOnline ? "theme-online" : "theme-offline";

                sb.AppendLine($"<div class='player-card {themeClass}'>");
                sb.AppendLine("<div class='player-heading'>");
                sb.AppendLine("<div>");
                sb.AppendLine($"<h2>{WebUtility.HtmlEncode(p.Name)}</h2>");
                sb.AppendLine($"<div class='player-meta'>BattleMetrics ID {WebUtility.HtmlEncode(p.BMId)}</div>");
                sb.AppendLine("</div>");
                
                var statusClass = isOnline ? "badge-online" : "badge-offline";
                var statusText = isOnline ? "Online" : "Offline";
                sb.AppendLine($"<span class='badge {statusClass}'>{statusText}</span>");
                sb.AppendLine("</div>");

                var lastS = p.Sessions.LastOrDefault();
                string lastConnectedStr = lastS != null ? lastS.ConnectTime.ToLocalTime().ToString("yyyy-MM-dd HH:mm") : "Never";
                string lastSeenStr = lastS != null ? (lastS.DisconnectTime?.ToLocalTime().ToString("yyyy-MM-dd HH:mm") ?? "Active Now") : "Never";

                sb.AppendLine("<div class='stat-grid'>");
                sb.AppendLine("<div class='stat-item'><div class='stat-label'>Last Connected</div><div class='stat-value'>" + lastConnectedStr + "</div></div>");
                sb.AppendLine("<div class='stat-item'><div class='stat-label'>Last Seen</div><div class='stat-value'>" + lastSeenStr + "</div></div>");
                sb.AppendLine("<div class='stat-item'><div class='stat-label'>Total Tracked Time</div><div class='stat-value'>" + $"{(int)totalTime.TotalHours}h {totalTime.Minutes}m" + "</div></div>");
                sb.AppendLine("<div class='stat-item'><div class='stat-label'>Last 7 Days</div><div class='stat-value'>" + $"{(int)past7Days.TotalHours}h {past7Days.Minutes}m" + "</div></div>");
                sb.AppendLine("<div class='stat-item'><div class='stat-label'>Session Count</div><div class='stat-value'>" + p.Sessions.Count + "</div></div>");
                sb.AppendLine("<div class='stat-item'><div class='stat-label'>Avg Session</div><div class='stat-value'>" + $"{(int)avgSessionMins} min" + "</div></div>");
                sb.AppendLine("</div>");

            // GitHub Style Grid Section
            sb.AppendLine("<div class='section-title'>12-WEEK ACTIVITY INTENSITY</div>");
            sb.AppendLine("<div class='grid-container'>");
            var startDate = now.Date.AddDays(-83); // 12 weeks
            for (int w = 0; w < 12; w++)
            {
                sb.AppendLine("<div class='grid-week'>");
                for (int d = 0; d < 7; d++)
                {
                    var cur = startDate.AddDays(w * 7 + d);
                    int mins = dailyActivity.ContainsKey(cur) ? dailyActivity[cur] : 0;
                    string lv = "";
                    if (mins > 0) lv = "lv1";
                    if (mins > 120) lv = "lv2";
                    if (mins > 300) lv = "lv3";
                    if (mins > 600) lv = "lv4";
                    sb.AppendLine($"<div class='grid-cell {lv}' title='{cur:yyyy-MM-dd}: {mins} min'></div>");
                }
                sb.AppendLine("</div>");
            }
            sb.AppendLine("</div>");

            // 24h Heatmap Section
            sb.AppendLine("<div class='section-title'>24H ACTIVITY FORECAST</div>");
            sb.AppendLine("<div class='hourly-wrap'>");
            sb.AppendLine("<div class='hourly-container'>");
            int maxH = hourActivity.Any() ? hourActivity.Max() : 0;
            for(int i=0; i<24; i++)
            {
                double hVal = maxH > 0 ? (double)hourActivity[i] / maxH * 100 : 5;
                string activeClass = hourActivity[i] > (maxH * 0.4) ? "active" : "";
                sb.AppendLine($"<div class='hour-bar {activeClass}' style='height:{hVal}%' title='{i:00}:00 - {hourActivity[i]} occurrences'></div>");
            }
            sb.AppendLine("</div>");
            sb.AppendLine("<div class='hour-labels'>");
            sb.AppendLine("<span>00:00</span><span>04:00</span><span>08:00</span><span>12:00</span><span>16:00</span><span>20:00</span><span>23:00</span>");
            sb.AppendLine("</div>");
            sb.AppendLine("</div>");

            // AI Insights Box
            int peakPlay = 0; int maxPlayVal = -1;
            int peakSleep = 0; int minPlayVal = int.MaxValue;
            for(int i=0; i<24; i++) {
                if (hourActivity[i] > maxPlayVal) { maxPlayVal = hourActivity[i]; peakPlay = i; }
                if (hourActivity[i] < minPlayVal) { minPlayVal = hourActivity[i]; peakSleep = i; }
            }

            sb.AppendLine("<div class='insight-box'>");
            sb.AppendLine("<div class='insight-item'><span class='insight-label'>Likely online</span><b>" + $"{peakPlay:00}:00 - {(peakPlay + 3) % 24:00}:00" + "</b></div>");
            sb.AppendLine("<div class='insight-item'><span class='insight-label'>Likely quiet</span><b>" + $"{peakSleep:00}:00 - {(peakSleep + 5) % 24:00}:00" + "</b></div>");
            if (p.Sessions.Count < 5) {
                sb.AppendLine("<div class='warning'><b>Data Confidence: LOW</b><br/>More sessions needed for accurate pattern recognition. Predictions currently represent early observations.</div>");
            } else {
                sb.AppendLine("<div style='color: #b8aaa0; font-size: 11px; margin-top: 10px;'>Forecast based on " + p.Sessions.Count + " recorded sessions.</div>");
            }
            sb.AppendLine("</div>");

            sb.AppendLine("</div>");
        }
    }
        
        sb.AppendLine("</div></body></html>");
        return sb.ToString();
    }

    private static string? _foundServerId;

    public static void StartPolling(string host, int port, string name)
    {
        _lastServerHost = host;
        _lastServerPort = port;
        _lastServerName = name;
        _foundServerId = null; // Reset to force new lookup

        _settings.LastHost = host;
        _settings.LastPort = port;
        _settings.LastServerName = name;
        SaveDB();

        // Poll every 2 minutes only if we have players
        _trackingTimer?.Dispose();
        if (_trackedPlayers.Count > 0)
        {
            _trackingTimer = new Timer(async _ => await PollOnceAsync(), null, 0, 120_000);
        }
        else
        {
            _trackingTimer = null;
        }
    }

    public static void StopPolling()
    {
        _trackingTimer?.Dispose();
        _trackingTimer = null;
    }

    public static async Task FetchOnlinePlayersNowAsync()
    {
        await PollOnceAsync();
    }

    private static async Task PollOnceAsync()
    {
        if (string.IsNullOrEmpty(_lastServerHost)) return;

        try
        {
            // 1. Get BM Server ID
            if (string.IsNullOrEmpty(_foundServerId))
            {
                StatusMessage = "Looking up server...";

                // --- SCHRITT A: Suche über IP-Adresse (Port ignorieren, da oft unterschiedlich) ---
                var searchUrlAddr = $"https://api.battlemetrics.com/servers?filter[address]={Uri.EscapeDataString(_lastServerHost)}&filter[game]=rust";

                using var responseAddr = await _http.GetAsync(searchUrlAddr);
                if (responseAddr.IsSuccessStatusCode)
                {
                    var resAddr = await responseAddr.Content.ReadAsStringAsync();
                    using var docAddr = JsonDocument.Parse(resAddr);
                    var dataArr = docAddr.RootElement.GetProperty("data");

                    foreach (var serverObj in dataArr.EnumerateArray())
                    {
                        var attr = serverObj.GetProperty("attributes");
                        var foundIp = attr.GetProperty("ip").GetString();
                        var foundName = attr.GetProperty("name").GetString() ?? "";

                        // CRITICAL: Wir nehmen den Server nur, wenn die IP EXAKT stimmt
                        if (foundIp == _lastServerHost)
                        {
                            if (attr.TryGetProperty("details", out var details) && details.TryGetProperty("rust_description", out var desc))
                            {
                                OnServerInfoUpdated?.Invoke(desc.GetString() ?? "");
                            }

                            // Wenn wir mehrere Server auf einer IP haben (Shared Hosting), 
                            // nehmen wir den, dessen Name am besten passt.
                            if (string.IsNullOrEmpty(_lastServerName) || foundName.Contains(_lastServerName, StringComparison.OrdinalIgnoreCase))
                            {
                                _foundServerId = serverObj.GetProperty("id").GetString();
                                break;
                            }
                        }
                    }
                }

                // --- SCHRITT B: Fallback über Namen (falls IP bei Battlemetrics anders gelistet ist) ---
                if (string.IsNullOrEmpty(_foundServerId) && !string.IsNullOrEmpty(_lastServerName))
                {
                    var searchUrlName = $"https://api.battlemetrics.com/servers?filter[game]=rust&filter[search]={Uri.EscapeDataString(_lastServerName)}";
                    using var responseName = await _http.GetAsync(searchUrlName);
                    if (responseName.IsSuccessStatusCode)
                    {
                        var resName = await responseName.Content.ReadAsStringAsync();
                        using var docName = JsonDocument.Parse(resName);
                        var dataArr = docName.RootElement.GetProperty("data");

                        foreach (var serverObj in dataArr.EnumerateArray())
                        {
                            var attr = serverObj.GetProperty("attributes");
                            var foundIp = attr.TryGetProperty("ip", out var vIp) ? vIp.GetString() : "";
                            var foundName = attr.GetProperty("name").GetString() ?? "";

                            // Wenn der Name exakt passt, nehmen wir die ID, auch wenn die IP leicht abweicht 
                            // (manche Server haben unterschiedliche IPs für Game und Websocket)
                            if (foundName.Equals(_lastServerName, StringComparison.OrdinalIgnoreCase))
                            {
                                _foundServerId = serverObj.GetProperty("id").GetString();
                                break;
                            }
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(_foundServerId))
            {
                StatusMessage = $"Server not found on Battlemetrics ({_lastServerHost}:{_lastServerPort})";
                OnOnlinePlayersUpdated?.Invoke();
                return;
            }

            // 2. Get Players using the ID we found
            StatusMessage = "Fetching players...";
            var reqUrl = $"https://api.battlemetrics.com/servers/{_foundServerId}?include=session";
            using var responsePlayers = await _http.GetAsync(reqUrl);
            if (!responsePlayers.IsSuccessStatusCode)
            {
                StatusMessage = $"Fetch Error: {(int)responsePlayers.StatusCode} {responsePlayers.ReasonPhrase}";
                OnOnlinePlayersUpdated?.Invoke();
                return;
            }

            var pRes = await responsePlayers.Content.ReadAsStringAsync();
            using var pDoc = JsonDocument.Parse(pRes);

            if (pDoc.RootElement.TryGetProperty("data", out var serverData))
            {
                var attr = serverData.GetProperty("attributes");
                if (attr.TryGetProperty("details", out var details) && details.TryGetProperty("rust_description", out var desc))
                {
                    OnServerInfoUpdated?.Invoke(desc.GetString() ?? "");
                }
            }
            
            var onlineList = new List<OnlinePlayerBM>();
            var currentlyOnlineInfo = new Dictionary<string, (DateTime start, string name)>();

            if (pDoc.RootElement.TryGetProperty("included", out var included))
            {
                foreach (var inc in included.EnumerateArray())
                {
                    string type = inc.TryGetProperty("type", out var tProp) ? tProp.GetString() ?? "" : "";
                    if (type == "session")
                    {
                        var attr = inc.GetProperty("attributes");
                        var name = attr.TryGetProperty("name", out var nProp) ? nProp.GetString() ?? "Unknown" : "Unknown";
                        var bmId = "";
                        
                        if (inc.TryGetProperty("relationships", out var rel) && 
                            rel.TryGetProperty("player", out var pRel) &&
                            pRel.TryGetProperty("data", out var pData))
                        {
                            bmId = pData.GetProperty("id").GetString() ?? "";
                        }
                        
                        if (string.IsNullOrEmpty(bmId)) continue;

                        int seconds = 0;
                        DateTime actualStart = DateTime.UtcNow;
                        if (attr.TryGetProperty("start", out var sProp) && sProp.ValueKind == JsonValueKind.String)
                        {
                            if (DateTimeOffset.TryParse(sProp.GetString(), out var start))
                            {
                                actualStart = start.UtcDateTime;
                                seconds = (int)(DateTimeOffset.UtcNow - start).TotalSeconds;
                            }
                        }

                        onlineList.Add(new OnlinePlayerBM
                        {
                            BMId = bmId,
                            Name = name,
                            SessionStartTimeUtc = actualStart,
                            Duration = TimeSpan.FromSeconds(Math.Max(0, seconds)),
                            IsTracked = _trackedPlayers.ContainsKey(bmId)
                        });
                        currentlyOnlineInfo[bmId] = (actualStart, name);
                    }
                }
            }

            if (onlineList.Count == 0)
            {
                StatusMessage = "No online players found on Battlemetrics.";
            }
            else
            {
                StatusMessage = "";
            }

            LastOnlinePlayers = onlineList.OrderByDescending(x => x.Duration).ToList();
            LastPullTime = DateTime.Now;
            OnOnlinePlayersUpdated?.Invoke();

            // 3. Update Tracking stats
            await UpdateTrackingStatsAsync(currentlyOnlineInfo);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Connection Error: {ex.Message}";
            OnOnlinePlayersUpdated?.Invoke();
        }
    }

    private static async Task UpdateTrackingStatsAsync(Dictionary<string, (DateTime start, string name)> currentlyOnlineInfo)
    {
        bool changed = false;
        var now = DateTime.UtcNow;

        foreach (var tp in _trackedPlayers.Values)
        {
            bool isOnline = currentlyOnlineInfo.TryGetValue(tp.BMId, out var info);
            var lastSession = tp.Sessions.LastOrDefault();

            if (isOnline)
            {
                // Update name if it was previously unknown or empty
                if (tp.Name == "Unknown Player" || string.IsNullOrEmpty(tp.Name))
                {
                    tp.Name = info.name;
                    changed = true;
                }

                var actualConnectTime = info.start;
                if (lastSession == null || lastSession.DisconnectTime.HasValue)
                {
                    // Newly connected or we just started tracking/opened the app
                    tp.Sessions.Add(new PlayerSession { ConnectTime = actualConnectTime, DisconnectTime = null });
                    Log($"[SESSION] {tp.Name} ({tp.BMId}) connected at {actualConnectTime:yyyy-MM-dd HH:mm:ss} UTC (detected at {now:HH:mm})");
                    changed = true;
                }
                else
                {
                    // If we have an open session, but the connect time is different (e.g. app was closed and they rejoined)
                    // BattleMetrics session ID would change, but here we track by server session.
                    // If the actualConnectTime is NEWER than our last recorded ConnectTime, they must have reconnected 
                    // while we were closed.
                    if (actualConnectTime > lastSession.ConnectTime.AddMinutes(5))
                    {
                        // They reconnected. Close old session at their last seen or roughly before this connect?
                        // For simplicity, we close the old one at actualConnectTime - 1 second and start new one.
                        lastSession.DisconnectTime = actualConnectTime.AddSeconds(-1);
                        tp.Sessions.Add(new PlayerSession { ConnectTime = actualConnectTime, DisconnectTime = null });
                        Log($"[SESSION] {tp.Name} reconnected (missed disconnect). New session start: {actualConnectTime:yyyy-MM-dd HH:mm:ss} UTC");
                        changed = true;
                    }
                    else if (Math.Abs((lastSession.ConnectTime - actualConnectTime).TotalMinutes) > 1)
                    {
                        // Small correction of start time
                        lastSession.ConnectTime = actualConnectTime;
                        changed = true;
                    }
                }
            }
            else
            {
                if (lastSession != null && !lastSession.DisconnectTime.HasValue)
                {
                    // Newly disconnected. Fetch actual last seen/stop time.
                    var actualDisconnectTime = await FetchLastSeenTimeAsync(tp.BMId);
                    if (actualDisconnectTime == DateTime.MinValue)
                    {
                        actualDisconnectTime = now;
                        Log($"[SESSION] {tp.Name} disconnected. API stop time fetch failed, using fallback: {now:yyyy-MM-dd HH:mm:ss} UTC");
                    }
                    else
                    {
                        Log($"[SESSION] {tp.Name} disconnected at {actualDisconnectTime:yyyy-MM-dd HH:mm:ss} UTC");
                    }
                    
                    lastSession.DisconnectTime = actualDisconnectTime;
                    changed = true;
                }
            }
        }

        if (changed)
        {
            SaveDB();
        }
    }

    private static async Task<DateTime> FetchLastSeenTimeAsync(string bmId)
    {
        if (string.IsNullOrEmpty(_foundServerId)) return DateTime.MinValue;

        try
        {
            // Fetch server-specific player information (free endpoint)
            var url = $"https://api.battlemetrics.com/players/{bmId}/servers/{_foundServerId}";
            var json = await _http.GetStringAsync(url);
            using var doc = JsonDocument.Parse(json);
            
            if (doc.RootElement.TryGetProperty("data", out var data) && 
                data.TryGetProperty("attributes", out var attr))
            {
                if (attr.TryGetProperty("lastSeen", out var stopProp) && stopProp.ValueKind == JsonValueKind.String)
                {
                    if (DateTimeOffset.TryParse(stopProp.GetString(), out var stop))
                    {
                        return stop.UtcDateTime;
                    }
                }
            }
        }
        catch { }
        return DateTime.MinValue;
    }
}
