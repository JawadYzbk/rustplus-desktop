using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using RustPlusDesk.Services;
using RustPlusDesk.Services.PlayerWipeTracker;

namespace RustPlusDesk.Views;

public partial class MainWindow
{
    private readonly PlayerWipeTrackerService _playerWipeTracker = new(
        new PlayerWipeTrackerStore(Path.Combine(RustPlusDesk.Services.Data.DataManager.AppDir, "player-wipes")),
        new PlayerWipeTrackerCapabilityService());

    private void StartPlayerWipeTrackerSession()
    {
        var profile = _vm?.Selected;
        var serverKey = GetServerKey();
        if (profile is null || string.IsNullOrWhiteSpace(serverKey))
            return;

        _playerWipeTracker.Enabled = Services.TrackingService.PlayerWipeTrackerEnabled;
        _playerWipeTracker.CloudBackupEnabled = Services.TrackingService.PlayerWipeTrackerCloudBackupEnabled;
        _playerWipeTracker.StartConnection(
            serverKey,
            (profile.WipeTime ?? profile.RustMapsWipeTime)?.ToUniversalTime(),
            profile.RustMapsMapId,
            _mySteamId);
    }

    private void StopPlayerWipeTrackerSession()
    {
        try { _playerWipeTracker.Disconnect(); } catch { }
    }

    private async Task DisposePlayerWipeTrackerAsync()
    {
        await _playerWipeTracker.DisposeAsync().ConfigureAwait(false);
    }

    private void ObservePlayerWipeTracker(RustPlusClientReal.TeamInfo team)
    {
        if (!Services.TrackingService.PlayerWipeTrackerEnabled || _playerWipeTracker.CurrentSessionId is null)
            return;

        _playerWipeTracker.Enabled = true;
        _playerWipeTracker.CloudBackupEnabled = Services.TrackingService.PlayerWipeTrackerCloudBackupEnabled;

        var classifier = new Services.Deaths.DeathLocationClassifier(
            BuildMonumentZones(), (x, y) => ResolveBaseAt(team, x, y), (x, y) => GetGridLabel(x, y));
        var timestamp = DateTime.UtcNow;
        foreach (var member in team.Members)
        {
            if (member.SteamId == 0)
                continue;

            var vm = TeamMembers.FirstOrDefault(item => item.SteamId == member.SteamId);
            var location = member.X.HasValue && member.Y.HasValue
                ? classifier.Classify(member.X, member.Y)
                : ("unknown", (string?)null);
            _playerWipeTracker.Observe(new PlayerObservation(
                member.SteamId,
                member.Name ?? vm?.Name ?? "(player)",
                timestamp,
                _playerWipeTracker.CurrentSessionId,
                true,
                true,
                member.Online,
                member.Dead,
                vm?.IsAfk ?? false,
                member.X,
                member.Y,
                location.Item1 switch
                {
                    "monument" => TrackerLocationType.Monument,
                    "base" => TrackerLocationType.Base,
                    "open" => TrackerLocationType.Open,
                    _ => TrackerLocationType.Unknown,
                },
                location.Item2,
                member.X.HasValue && member.Y.HasValue ? classifier.Grid(member.X, member.Y) : null,
                member.SpawnTime,
                member.DeathTime));
        }
    }

    private async Task RefreshPlayerWipeTrackerCapabilitiesAsync()
    {
        try
        {
            var client = new LaravelPlayerWipeTrackerClient();
            using var response = await client.GetBootstrapAsync().ConfigureAwait(false);
            if (response is not null)
                _playerWipeTracker.UpdateCapabilities(response.RootElement);
        }
        catch
        {
            // Capability outages must not affect Rust+ polling.
        }
    }

    private void BtnWipeTracker_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        try
        {
            var window = new RustPlusDesk.Views.Windows.PlayerWipeTrackerWindow(
                _playerWipeTracker,
                _mySteamId,
                ImgMap.Source,
                _worldSizeS,
                _worldRectPx)
            {
                Owner = this,
            };
            window.Show();
        }
        catch { }
    }
}
