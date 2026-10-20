using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Updates;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the update service: the running version, the newest published release and the way to get it. The
/// preview asks its own repository, because the release that replaces the preview is the preview's release.
/// </summary>
public sealed partial class UpdateSettingsViewModel(Func<Task<ReleaseInfo?>>? latest) : ObservableObject
{
    /// <summary>The name of the installer asset the release carries.</summary>
    public const string InstallerAssetName = "RustPlusDesk-Avalonia-Setup.exe";

    [ObservableProperty] private string status = "";
    [ObservableProperty] private bool isChecking;
    [ObservableProperty] private bool isUpdateAvailable;
    [ObservableProperty] private string latestVersion = "";

    /// <summary>The running version of the app.</summary>
    public string CurrentVersion => AppVersion.Short(AppVersion.Raw()).ToString();

    public string CurrentVersionText => $"{CurrentVersion} ({AppVersion.Raw()})";

    /// <summary>The page with the releases, opened when the user wants the new version.</summary>
    public string ReleasesPageUrl { get; set; } = "";

    /// <summary>Opens a link in the user's browser; the view owns it, because only a window can.</summary>
    public Func<string, Task<bool>>? OpenUrl { get; set; }

    /// <summary>Reference the "open the releases" action of the update card.</summary>
    [RelayCommand]
    private async Task OpenReleasesAsync()
    {
        if (OpenUrl is not { } open || ReleasesPageUrl.Length == 0) return;
        await open(ReleasesPageUrl);
    }

    /// <summary>Reference CheckForUpdates: the newest release is read and the user is told whether it is newer.</summary>
    [RelayCommand]
    private async Task CheckAsync(CancellationToken ct)
    {
        if (latest is not { } read) { Status = "The update service is not available."; return; }
        IsChecking = true;
        Status = "Checking for a new version…";
        try
        {
            var release = await read();
            if (release is null)
            {
                Status = "No release could be read. Check the connection and try again.";
                IsUpdateAvailable = false;
                return;
            }
            LatestVersion = release.Version.ToString();
            IsUpdateAvailable = AppVersion.IsNewer(AppVersion.Short(CurrentVersion), release.Version);
            Status = IsUpdateAvailable
                ? $"Version {release.Version} is available. This build is {CurrentVersion}."
                : $"This build is up to date ({CurrentVersion}).";
        }
        catch (Exception ex)
        {
            Status = $"The update check failed: {ex.Message}";
            IsUpdateAvailable = false;
        }
        finally { IsChecking = false; }
    }
}
