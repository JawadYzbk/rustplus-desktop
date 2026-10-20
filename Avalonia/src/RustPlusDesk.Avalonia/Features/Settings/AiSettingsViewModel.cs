using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the AI companion's settings: which provider answers, which model, and the key that pays for it.
///
/// The key is written to a private file and never sent anywhere but the provider, and the card says so.
/// </summary>
public sealed partial class AiSettingsViewModel(IAiSettingsStore store) : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSetUp))]
    private string provider = AiProviders.OpenAi;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSetUp))]
    private string model = "gpt-4o-mini";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSetUp))]
    private string apiKey = "";
    [ObservableProperty] private string status = "";

    public IReadOnlyList<string> Providers { get; } = [AiProviders.OpenAi, AiProviders.OpenRouter];

    /// <summary>Whether the companion is set up (reference AiCompanionStore.HasKey).</summary>
    public bool IsSetUp => ApiKey.Trim().Length > 0;

    /// <summary>Reads the saved settings.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        var settings = await store.LoadAsync(ct);
        Provider = settings.Provider;
        Model = settings.Model;
        ApiKey = settings.ApiKey;
        OnPropertyChanged(nameof(IsSetUp));
    }

    /// <summary>Saves the settings; the key goes into the private file.</summary>
    [RelayCommand]
    private async Task SaveAsync(CancellationToken ct)
    {
        try
        {
            await store.SaveAsync(new AiSettings { Provider = Provider, Model = Model.Trim(), ApiKey = ApiKey.Trim() }, ct);
            Status = "The AI companion is saved on this machine only.";
        }
        catch (Exception ex)
        {
            Status = $"The settings could not be saved: {ex.Message}";
        }
    }
}
