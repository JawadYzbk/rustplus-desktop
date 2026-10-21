using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>One provider in the picker: the id that is stored, and the name that is shown.</summary>
/// <param name="Id">The provider's id, as it is written to the settings file.</param>
/// <param name="Name">The name shown in the picker.</param>
public sealed record AiProviderChoice(string Id, string Name);

/// <summary>
/// Reference the AI companion's settings: which provider answers, which model it is asked, and the key that pays for it.
///
/// The key is written to a private file and never sent anywhere but the provider, and the card says so.
///
/// One slot per provider rather than one slot for the lot: people keep more than one key and switch between them, and a
/// single slot meant the other account's key was quietly overwritten. The card shows which of the other providers are
/// set up too, because that is the whole reason there is a slot each.
/// </summary>
public sealed partial class AiSettingsViewModel(IAiSettingsStore store) : ObservableObject
{
    /// <summary>The settings as they are stored, and the only copy that is ever written back.</summary>
    private AiSettings settings = new();

    /// <summary>Which provider answers.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ProviderName))]
    [NotifyPropertyChangedFor(nameof(ProviderNote))]
    [NotifyPropertyChangedFor(nameof(ProviderKeyUrl))]
    [NotifyPropertyChangedFor(nameof(KeyState))]
    [NotifyPropertyChangedFor(nameof(IsSetUp))]
    [NotifyPropertyChangedFor(nameof(HasKey))]
    [NotifyPropertyChangedFor(nameof(OtherKeysStored))]
    [NotifyPropertyChangedFor(nameof(SuggestedModels))]
    [NotifyPropertyChangedFor(nameof(IsOpenRouter))]
    [NotifyPropertyChangedFor(nameof(DefaultModel))]
    private string provider = AiProviders.OpenAi;

    /// <summary>The model the chosen provider is asked for, as typed.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanResetModel))]
    private string model = "";

    /// <summary>The key for the chosen provider, as typed. It is never shown again once it is saved.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(KeyLooksValid))]
    private string apiKey = "";

    /// <summary>What the card says after a save, a removal or a failure.</summary>
    [ObservableProperty] private string status = "";

    /// <summary>The providers, in the reference's order.</summary>
    public IReadOnlyList<AiProviderChoice> Providers { get; } =
        AiProviders.All.Select(id => new AiProviderChoice(id, AiProviders.DisplayName(id))).ToArray();

    /// <summary>
    /// The provider the picker shows as chosen.
    ///
    /// The picker works on the named choice rather than on the id, so the list can say "OpenAI - GPT" while the file
    /// keeps "openai".
    /// </summary>
    public AiProviderChoice SelectedProvider
    {
        get => Providers.FirstOrDefault(choice => choice.Id == Provider) ?? Providers[0];
        set
        {
            if (value is not null) Provider = value.Id;
        }
    }

    /// <summary>The name of the chosen provider.</summary>
    public string ProviderName => AiProviders.DisplayName(Provider);

    /// <summary>The model the chosen provider falls back to when none has been typed.</summary>
    public string DefaultModel => AiProviders.DefaultModel(Provider);

    /// <summary>Whether the chosen provider is OpenRouter, which is the one with a catalogue of its own.</summary>
    public bool IsOpenRouter => Provider == AiProviders.OpenRouter;

    /// <summary>Whether a key is stored for the chosen provider.</summary>
    public bool IsSetUp => settings.HasKeyFor(Provider);

    /// <summary>Whether a key is stored for the chosen provider, for the remove button.</summary>
    public bool HasKey => IsSetUp;

    /// <summary>The other providers that are set up too, as a sentence under the key.</summary>
    public string OtherKeysStored
    {
        get
        {
            var others = AiProviders.All
                .Where(other => other != Provider && settings.HasKeyFor(other))
                .Select(AiProviders.DisplayName)
                .ToList();
            return others.Count == 0
                ? ""
                : string.Format(Loc.Text("AiCompanionKeyAlsoStored", "Also stored: {0}."), string.Join(", ", others));
        }
    }

    /// <summary>What the card says about the key: stored, or missing, and for which provider.</summary>
    public string KeyState => IsSetUp
        ? string.Format(Loc.Text("AiCompanionKeyStoredFor", "A {0} key is stored on this PC."), ProviderName)
        : string.Format(Loc.Text("AiCompanionKeyMissingFor", "No {0} key yet - this provider cannot send anything without one."), ProviderName);

    /// <summary>Whether the typed key looks like a key at all, so a paste error is caught here.</summary>
    public bool KeyLooksValid => string.IsNullOrWhiteSpace(ApiKey) || AiProviders.LooksLikeKey(Provider, ApiKey);

    /// <summary>Whether the typed model can be put back to the provider's default.</summary>
    public bool CanResetModel => !string.IsNullOrWhiteSpace(Model);

    /// <summary>One-click model choices. The text box stays the source of truth.</summary>
    public IReadOnlyList<string> SuggestedModels => AiProviders.SuggestedModels(Provider);

    /// <summary>Where the key for this provider comes from.</summary>
    public string ProviderKeyUrl => AiProviders.KeyUrl(Provider);

    /// <summary>
    /// What this provider is, in the card's own words.
    ///
    /// The reference tells the player what the provider can and cannot do - whether it takes a recording, whether it has
    /// a voice - and what it costs. This port answers in text only for now, so it says that instead of describing a
    /// recording it cannot take.
    /// </summary>
    public string ProviderNote => string.Format(Loc.Text("AiCompanionProviderText", "{0} answers in text on this build. The key is sent to the provider and nowhere else."), ProviderName)
        + " " + (Provider switch
        {
            AiProviders.Gemini => Loc.Text("AiCompanionBillingFree",
                "Google gives the API a free allowance, so a key from AI Studio works without paying. Busy sessions can still hit the per-minute limit."),
            AiProviders.OpenRouter => Loc.Text("AiCompanionBillingOpenRouter",
                "One key for many models with pay-as-you-go credit. Prices vary per model - cheap ones cost a fraction of GPT per question."),
            _ => Loc.Text("AiCompanionBillingPrepaid",
                "This API is billed separately from any chat subscription and needs its own prepaid credit."),
        });

    // ---- OpenRouter's own request options --------------------------------------------------

    /// <summary>How deeply the reasoning models think before answering.</summary>
    [ObservableProperty] private string reasoningEffort = "auto";

    /// <summary>The ceiling on the reasoning tokens, zero for the model's own default.</summary>
    [ObservableProperty] private int reasoningMaxTokens;

    /// <summary>How creative the answer is allowed to be.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TemperatureLabel))]
    private double temperature = 0.7;

    /// <summary>How long the answer may be.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(MaxTokensSlider))]
    private int maxTokens = 700;

    /// <summary>Whether OpenRouter may try another host when the first one is down.</summary>
    [ObservableProperty] private bool allowFallbacks = true;

    /// <summary>Whether the request is kept away from the providers that log what was sent.</summary>
    [ObservableProperty] private bool strictPrivacy;

    /// <summary>The temperature, as the slider's label shows it.</summary>
    public string TemperatureLabel => $"Temperature - {Temperature:0.00}";

    /// <summary>
    /// The maximum output length, as the slider shows it.
    ///
    /// The slider works in fractions of a token and the setting is a whole number of them, so the two are bridged here
    /// rather than in the markup.
    /// </summary>
    public double MaxTokensSlider
    {
        get => MaxTokens;
        set
        {
            int rounded = (int)Math.Round(value);
            if (rounded != MaxTokens) MaxTokens = rounded;
        }
    }

    /// <summary>The reasoning efforts the picker offers.</summary>
    public IReadOnlyList<string> ReasoningEfforts { get; } = ["auto", "low", "medium", "high", "none"];

    /// <summary>The thinking-token ceilings the picker offers.</summary>
    public IReadOnlyList<int> ThinkingTokenChoices { get; } = [0, 1024, 2048, 4096, 8192, 16384, 32768];

    /// <summary>
    /// Opens the model browser.
    ///
    /// The window belongs to the shell, so the card only asks for it - the same way the other windows in this app are
    /// opened by the shell rather than by a page.
    /// </summary>
    public Action? BrowseModels { get; set; }

    /// <summary>Reads the saved settings.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        settings = await store.LoadAsync(ct);
        Provider = settings.Provider;
        Model = settings.ModelFor(Provider);
        ApiKey = settings.KeyFor(Provider);

        var openRouter = settings.OpenRouter;
        ReasoningEffort = string.IsNullOrWhiteSpace(openRouter.ReasoningEffort) ? "auto" : openRouter.ReasoningEffort;
        ReasoningMaxTokens = openRouter.ReasoningMaxTokens;
        Temperature = openRouter.Temperature;
        MaxTokens = openRouter.MaxTokens <= 0 ? 700 : openRouter.MaxTokens;
        AllowFallbacks = openRouter.AllowFallbacks;
        StrictPrivacy = string.Equals(openRouter.DataCollection, "deny", StringComparison.OrdinalIgnoreCase);

        OnPropertyChanged(nameof(IsSetUp));
        OnPropertyChanged(nameof(HasKey));
        OnPropertyChanged(nameof(KeyState));
        OnPropertyChanged(nameof(OtherKeysStored));
        OnPropertyChanged(nameof(SuggestedModels));
        OnPropertyChanged(nameof(DefaultModel));
    }

    /// <summary>Switches provider: the model and the key are per provider, so each brings its own.</summary>
    partial void OnProviderChanged(string value)
    {
        Model = settings.ModelFor(value);
        ApiKey = settings.KeyFor(value);
        OnPropertyChanged(nameof(SelectedProvider));
        OnPropertyChanged(nameof(CanResetModel));
        OnPropertyChanged(nameof(KeyLooksValid));
    }

    /// <summary>Asks the shell for the model browser.</summary>
    [RelayCommand]
    private void Browse() => BrowseModels?.Invoke();

    /// <summary>
    /// Reference CmbAiModel_SelectionChanged: a model picked from the list fills the box, which stays the source of
    /// truth.
    /// </summary>
    [RelayCommand]
    private void UseSuggestedModel(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        Model = name;
        Status = "";
    }

    /// <summary>Reference BtnAiModelReset_Click: the provider's own default is used again.</summary>
    [RelayCommand]
    private void ResetModel()
    {
        Model = "";
        Status = string.Format(Loc.Text("AiCompanionModelNote", "The model is now {0}."), DefaultModel);
    }

    /// <summary>
    /// Reference BtnSaveAiKey_Click: the key is stored for this provider, and the model with it.
    ///
    /// The other providers' keys and models are left exactly as they were.
    /// </summary>
    [RelayCommand]
    private async Task SaveAsync(CancellationToken ct)
    {
        try
        {
            AiSettings updated = settings
                .WithProvider(Provider)
                .WithModel(Provider, Model)
                .WithKey(Provider, ApiKey)
                .WithOpenRouter(new OpenRouterSettings
                {
                    ReasoningEffort = ReasoningEffort,
                    ReasoningMaxTokens = ReasoningMaxTokens,
                    Temperature = Temperature,
                    MaxTokens = MaxTokens,
                    AllowFallbacks = AllowFallbacks,
                    DataCollection = StrictPrivacy ? "deny" : "allow",
                });

            await store.SaveAsync(updated, ct);

            settings = updated;
            ApiKey = settings.KeyFor(Provider);
            Status = Loc.Text("AiCompanionKeyStored", "The key is saved on this machine only.");
            OnPropertyChanged(nameof(IsSetUp));
            OnPropertyChanged(nameof(HasKey));
            OnPropertyChanged(nameof(KeyState));
            OnPropertyChanged(nameof(OtherKeysStored));
        }
        catch (Exception ex)
        {
            Status = Loc.Text("AiCompanionKeyRejectedTitle", "The key was not saved.") + " " + ex.Message;
        }
    }

    /// <summary>Reference BtnRemoveAiKey_Click: the key for this provider is removed and nothing else is touched.</summary>
    [RelayCommand]
    private async Task RemoveKeyAsync(CancellationToken ct)
    {
        try
        {
            AiSettings updated = settings.WithKey(Provider, "");
            await store.SaveAsync(updated, ct);

            settings = updated;
            ApiKey = "";
            Status = Loc.Text("AiCompanionKeyMissing", "The key was removed from this machine.");
            OnPropertyChanged(nameof(IsSetUp));
            OnPropertyChanged(nameof(HasKey));
            OnPropertyChanged(nameof(KeyState));
            OnPropertyChanged(nameof(OtherKeysStored));
        }
        catch (Exception ex)
        {
            Status = Loc.Text("AiCompanionKeyRejectedTitle", "The key was not saved.") + " " + ex.Message;
        }
    }

    /// <summary>
    /// A model picked in the browser fills the box, which is still the source of truth.
    /// </summary>
    public void ApplyModel(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        Model = name.Trim();
        Status = "";
    }
}
