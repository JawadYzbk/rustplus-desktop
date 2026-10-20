using CommunityToolkit.Mvvm.ComponentModel;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the language picker: the app reads its strings in the language the user chose, and follows the
/// operating system's language until they choose one.
/// </summary>
public sealed partial class LanguageSettingsViewModel(IAppSettingsStore? settings) : ObservableObject
{
    /// <summary>Every language the app ships (reference AppLanguages.All).</summary>
    public IReadOnlyList<AppLanguage> Languages { get; } = AppLanguages.All;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CurrentLanguage))]
    private string selectedCode = "";

    /// <summary>The language in words, or the one the operating system reports.</summary>
    public string CurrentLanguage
    {
        get
        {
            if (SelectedCode.Length > 0)
                return Languages.FirstOrDefault(language => language.Code == SelectedCode)?.Name ?? SelectedCode;
            string? resolved = AppLanguages.Resolve(System.Globalization.CultureInfo.CurrentUICulture);
            return Languages.FirstOrDefault(language => language.Code == resolved)?.Name ?? "System language";
        }
    }

    /// <summary>Reads the saved language and puts the app in it.</summary>
    public void Load()
    {
        SelectedCode = settings?.Get(SettingKeys.SelectedLanguage, "") ?? "";
        Apply();
    }

    partial void OnSelectedCodeChanged(string value)
    {
        if (loading) return;
        settings?.Set(SettingKeys.SelectedLanguage, value);
        Apply();
    }

    private bool loading;

    /// <summary>Called when the language changed, so the labels can be read again.</summary>
    public Action? LanguageChanged { get; set; }

    /// <summary>Puts the app in the chosen language, or in the operating system's when none is chosen.</summary>
    public void Apply()
    {
        loading = true;
        var culture = AppLanguages.CultureFor(SelectedCode)
            ?? AppLanguages.CultureFor(AppLanguages.Resolve(System.Globalization.CultureInfo.CurrentUICulture));
        Loc.UseCulture(culture);
        loading = false;
        OnPropertyChanged(nameof(CurrentLanguage));
        LanguageChanged?.Invoke();
    }
}
