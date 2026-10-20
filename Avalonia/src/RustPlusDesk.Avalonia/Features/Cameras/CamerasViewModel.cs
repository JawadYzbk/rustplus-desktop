using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Cameras;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Features.Cameras;

/// <summary>
/// Reference CameraSession: a camera the user added, as the tab shows it - the picture, the kind of camera and
/// whether it is live right now.
/// </summary>
public sealed partial class CameraRow : ObservableObject
{
    public CameraRow(CameraEntry entry)
    {
        Id = entry.Id;
        Name = entry.DisplayName;
    }

    /// <summary>The camera's identifier, the way the server knows it.</summary>
    public string Id { get; }

    /// <summary>The name the user gave the camera.</summary>
    public string Name { get; set; }

    /// <summary>Reference GrabCameraSnapshotAsync: the last picture of the camera.</summary>
    [ObservableProperty] private byte[]? thumbnail;

    /// <summary>The picture as the view shows it.</summary>
    public global::Avalonia.Media.Imaging.Bitmap? Picture { get; private set; }

    partial void OnThumbnailChanged(byte[]? value)
    {
        if (value is null) { Picture = null; return; }
        using var stream = new System.IO.MemoryStream(value);
        try { Picture = new global::Avalonia.Media.Imaging.Bitmap(stream); }
        catch { Picture = null; }
    }

    /// <summary>Reference CamDisplayName: the kind of camera, once it was opened.</summary>
    [ObservableProperty] private string kind = "";

    /// <summary>Whether the camera answered.</summary>
    [ObservableProperty] private bool isLive;
}

/// <summary>
/// Reference MainWindow.Camera: the cameras the user added for the server.
///
/// The reference grabs a picture of one camera at a time, every few seconds, so the tiles show what the cameras see
/// without asking the server for all of them at once. It opens the camera in a window of its own when one is clicked.
/// </summary>
public sealed partial class CamerasViewModel(ICameraStore store, ICameraSnapshotSource snapshots) : ObservableObject
{
    private string profilesPath = "";
    private string profileKey = "";
    private IReadOnlyList<CameraRow> rows = [];
    private int next;

    /// <summary>Reference CamerasTabContent: the cameras the user added.</summary>
    public IReadOnlyList<CameraRow> Rows => rows;

    /// <summary>Reference AddCamPopup: the camera the user is adding.</summary>
    [ObservableProperty] private string newId = "";

    /// <summary>Reference AddCamName: the name the user gives the camera.</summary>
    [ObservableProperty] private string newName = "";

    /// <summary>Reference AddCamErr: what is wrong with the camera the user typed in.</summary>
    [ObservableProperty] private string error = "";

    /// <summary>Reference AddCamPopup: the form where a camera is added is open.</summary>
    [ObservableProperty] private bool isAdding;

    /// <summary>Whether there is nothing to show yet.</summary>
    public bool IsEmpty => rows.Count == 0;

    /// <summary>Reference MainWindow.Camera: the cameras are read for the selected server.</summary>
    public async Task LoadAsync(string path, string key, CancellationToken ct = default)
    {
        profilesPath = path;
        profileKey = key;
        var cameras = await store.LoadAsync(path, key, ct).ConfigureAwait(false);
        rows = cameras.Select(camera => new CameraRow(camera)).ToArray();
        Refresh();
    }

    /// <summary>Reference InitCameraUi: the form where a camera is added opens.</summary>
    [RelayCommand]
    private void OpenAdd()
    {
        IsAdding = true;
        Error = "";
    }

    /// <summary>Reference CommitAddCamera: the camera the user typed in is added.</summary>
    [RelayCommand]
    private async Task AddAsync()
    {
        string id = NewId.Trim();
        if (id.Length == 0)
        {
            Error = "Please enter a camera identifier.";
            return;
        }
        if (rows.Any(row => string.Equals(row.Id, id, StringComparison.OrdinalIgnoreCase)))
        {
            Error = "That camera is already added.";
            return;
        }

        await SaveAsync(rows.Select(row => new CameraEntry(row.Id, row.Name))
            .Append(new CameraEntry(id, NewName.Trim())).ToArray()).ConfigureAwait(false);
        NewId = "";
        NewName = "";
        Error = "";
        IsAdding = false;
        await LoadAsync(profilesPath, profileKey).ConfigureAwait(false);
    }

    /// <summary>Reference CommitEditCamera: the camera's identifier and name are changed.</summary>
    [RelayCommand]
    private async Task EditAsync(CameraRow row)
    {
        string id = row.Id.Trim();
        if (id.Length == 0)
        {
            Error = "Please enter a camera identifier.";
            return;
        }
        if (rows.Any(other => !ReferenceEquals(other, row) && string.Equals(other.Id, id, StringComparison.OrdinalIgnoreCase)))
        {
            Error = "That camera is already added.";
            return;
        }

        var cameras = rows.Select(other => ReferenceEquals(other, row)
            ? new CameraEntry(id, other.Name)
            : new CameraEntry(other.Id, other.Name)).ToArray();
        await SaveAsync(cameras).ConfigureAwait(false);
        Error = "";
        await LoadAsync(profilesPath, profileKey).ConfigureAwait(false);
    }

    /// <summary>Reference CommitDeleteCamera: the camera is no longer added.</summary>
    [RelayCommand]
    private async Task RemoveAsync(CameraRow row)
    {
        await SaveAsync(rows.Where(other => !ReferenceEquals(other, row))
            .Select(other => new CameraEntry(other.Id, other.Name)).ToArray()).ConfigureAwait(false);
        Error = "";
        await LoadAsync(profilesPath, profileKey).ConfigureAwait(false);
    }

    /// <summary>
    /// Reference CamThumbTimer_Tick: one camera at a time is asked for a picture, so the server is not asked for all of
    /// them at once.
    /// </summary>
    public async Task GrabNextAsync(CancellationToken ct = default)
    {
        if (rows.Count == 0) return;
        if (next >= rows.Count) next = 0;
        var row = rows[next++];
        var picture = await snapshots.GrabAsync(profilesPath, profileKey, row.Id, ct).ConfigureAwait(false);
        if (picture is null)
        {
            row.IsLive = false;
            return;
        }
        row.Thumbnail = picture;
        row.IsLive = true;
    }

    private async Task SaveAsync(IReadOnlyList<CameraEntry> cameras) =>
        await store.SaveAsync(profilesPath, profileKey, cameras, CancellationToken.None).ConfigureAwait(false);

    private void Refresh()
    {
        OnPropertyChanged(nameof(Rows));
        OnPropertyChanged(nameof(IsEmpty));
    }
}
