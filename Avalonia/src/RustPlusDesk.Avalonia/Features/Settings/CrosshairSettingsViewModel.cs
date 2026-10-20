using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Crosshair;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the crosshair settings: the style, the custom image and the switch that shows the crosshair over the
/// game. The style is kept in the app's settings, so it is there again the next time the app starts.
/// </summary>
public sealed partial class CrosshairSettingsViewModel(IAppSettingsStore? settings) : ObservableObject
{
    /// <summary>While the saved settings are read, nothing is written back.</summary>
    private bool loading;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StyleLabel))]
    private CrosshairStyle style = CrosshairStyle.GreenDot;

    /// <summary>The image the custom style shows.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StyleLabel))]
    private string? customImagePath;

    /// <summary>Whether the crosshair is shown over the game.</summary>
    [ObservableProperty]
    private bool enabled;

    public string StyleLabel => CustomImagePath is { Length: > 0 } && Style == CrosshairStyle.Custom
        ? $"Custom image ({Path.GetFileName(CustomImagePath)})"
        : CrosshairCatalogue.Label(Style);

    public IReadOnlyList<CrosshairStyle> Styles { get; } = CrosshairCatalogue.All;

    /// <summary>Raised when the crosshair should be shown or hidden.</summary>
    public event Action<bool>? VisibilityChanged;

    /// <summary>Reads the saved crosshair settings.</summary>
    public void Load()
    {
        if (settings is null) return;
        loading = true;
        Style = Enum.TryParse(settings.Get(SettingKeys.CrosshairStyle, CrosshairStyle.GreenDot.ToString()), out CrosshairStyle parsed)
            ? parsed : CrosshairStyle.GreenDot;
        CustomImagePath = settings.Get<string?>(SettingKeys.CrosshairCustomImagePath, null);
        Enabled = settings.Get(SettingKeys.CrosshairEnabled, false);
        loading = false;
    }

    partial void OnStyleChanged(CrosshairStyle value) { if (loading) return; Save(); VisibilityChanged?.Invoke(Enabled); }
    partial void OnCustomImagePathChanged(string? value) { if (loading) return; Save(); VisibilityChanged?.Invoke(Enabled); }
    partial void OnEnabledChanged(bool value) { if (loading) return; Save(); VisibilityChanged?.Invoke(value); }

    /// <summary>Reference the custom image picker: the file the crosshair shows.</summary>
    [RelayCommand]
    private async Task PickImageAsync(Func<Task<string?>>? pick = null)
    {
        if (pick is null) return;
        string? path = await pick();
        if (!string.IsNullOrWhiteSpace(path))
        {
            CustomImagePath = path;
            Style = CrosshairStyle.Custom;
        }
    }

    private void Save()
    {
        if (settings is null) return;
        settings.Set(SettingKeys.CrosshairStyle, Style.ToString());
        settings.Set(SettingKeys.CrosshairCustomImagePath, CustomImagePath);
        settings.Set(SettingKeys.CrosshairEnabled, Enabled);
    }
}
