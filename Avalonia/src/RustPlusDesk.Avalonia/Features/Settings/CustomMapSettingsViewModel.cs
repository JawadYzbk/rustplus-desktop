using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference "Custom HD Map Image" card (AppSettingsOverlay TxtCustomMapUrl / BtnApplyCustomMapUrl / BtnClearCustomMapUrl):
/// the selected profile's CustomMapUrl with Apply and Clear. Applying validates the URL, saves it, drops the cached
/// custom image for a new URL and reloads the map.
/// </summary>
public sealed partial class CustomMapSettingsViewModel(ICustomMapUrlStore store, IMapCache? cache, Func<(string Path, ServerSummary Server)?> selected,
    Func<Task> reload) : ObservableObject
{
    [ObservableProperty] private string url = "";
    [ObservableProperty] private string status = CustomMapUrl.NoServerStatus;
    [ObservableProperty] private bool isError;

    /// <summary>Reference settings load: the selected profile's URL and the matching status.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        IsError = false;
        if (selected() is not { } target)
        {
            Url = "";
            Status = CustomMapUrl.NoServerStatus;
            return;
        }
        string? current;
        try { current = await store.GetAsync(target.Path, target.Server.ProfileKey, ct); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException) { current = null; }
        Url = current ?? "";
        Status = current is null ? CustomMapUrl.NoneStatus : CustomMapUrl.ActiveStatus;
    }

    [RelayCommand]
    private async Task ApplyAsync(CancellationToken ct)
    {
        if (selected() is not { } target)
        {
            (Status, IsError) = (CustomMapUrl.NoServerStatus, true);
            return;
        }
        if (!CustomMapUrl.TryNormalize(Url, out string? normalized))
        {
            (Status, IsError) = (CustomMapUrl.InvalidMessage, true);
            return;
        }
        try { await store.SetAsync(target.Path, target.Server.ProfileKey, normalized, ct); }
        catch (ProfileChangedException) { (Status, IsError) = ("The profiles file changed. Refresh the server list and try again.", true); return; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        { (Status, IsError) = ("Unable to save. Check file permissions and available disk space.", true); return; }
        Url = normalized ?? "";
        IsError = false;
        if (normalized is not null)
        {
            cache?.DeleteCustom(target.Server.Host, target.Server.Port);
            Status = CustomMapUrl.UpdatedStatus;
        }
        else Status = CustomMapUrl.ClearedStatus;
        await reload();
    }

    /// <summary>Reference BtnClearCustomMapUrl_Click: empties the box and applies.</summary>
    [RelayCommand]
    private Task ClearAsync(CancellationToken ct)
    {
        Url = "";
        return ApplyAsync(ct);
    }
}
