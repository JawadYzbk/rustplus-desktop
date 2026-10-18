using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Alerts;

/// <summary>
/// Reference Discord webhook for chat alerts (free tier): the selected profile's URL, optional @everyone/@here pings,
/// TTS and "Discord only". While enabled and signed in to the cloud, every automated alert is also posted as
/// "**[{server}]** {message}"; "Discord only" keeps it out of in-game chat.
/// </summary>
public sealed partial class DiscordWebhookViewModel(IDiscordWebhookStore store, IDiscordWebhookSender sender, ICloudSignInState? signIn,
    Func<(string Path, string ProfileKey, string? Name)?> profile) : ObservableObject
{
    private DiscordWebhookSettings settings = new();
    private int version;

    public Action<string>? Log { get; set; }
    /// <summary>Raised when Enabled changes (the alert count and master follow it).</summary>
    public event Action? EnabledChanged;

    /// <summary>Reference IsCloudConnected: the webhook only sends (and can be switched on) while signed in.</summary>
    [ObservableProperty] private bool isSignedIn;
    /// <summary>The URL being edited; Save stores it.</summary>
    [ObservableProperty] private string urlDraft = "";
    [ObservableProperty] private string status = "";

    public DiscordWebhookSettings Settings => settings;
    public bool HasProfile => profile() is not null;
    public bool Enabled => settings.Enabled;
    public bool Exclusive { get => settings.Exclusive; set => Change(settings with { Exclusive = value }); }
    public bool Tts { get => settings.Tts; set => Change(settings with { Tts = value }); }
    public bool MentionEveryone { get => settings.MentionsEveryone; set => Change(settings.WithMention(DiscordWebhookSettings.Everyone, value)); }
    public bool MentionHere { get => settings.MentionsHere; set => Change(settings.WithMention(DiscordWebhookSettings.Here, value)); }
    public string SignInNote => IsSignedIn ? "" : "Sign in to Rust+ Desktop Cloud to send chat alerts to Discord.";
    partial void OnIsSignedInChanged(bool value) => OnPropertyChanged(nameof(SignInNote));

    /// <summary>Reads the selected profile's webhook and the cloud sign-in (server selection, app start).</summary>
    public async Task LoadAsync()
    {
        int current = ++version;
        var loaded = new DiscordWebhookSettings();
        if (profile() is { } target)
        {
            try { loaded = await store.GetAsync(target.Path, target.ProfileKey, CancellationToken.None); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException) { loaded = new(); }
        }
        bool signedIn = signIn is not null && await signIn.IsSignedInAsync(CancellationToken.None);
        if (current != version) return;
        IsSignedIn = signedIn;
        Apply(loaded);
        UrlDraft = loaded.Url;
        Status = "";
    }

    /// <summary>Stores the URL; a URL turns the webhook on, an empty one off (reference setter).</summary>
    [RelayCommand]
    private async Task SaveUrlAsync()
    {
        string url = UrlDraft.Trim();
        if (url.Length > 0 && !(Uri.TryCreate(url, UriKind.Absolute, out var uri) && (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp)))
        {
            Status = "Enter the webhook URL Discord shows under Server Settings > Integrations > Webhooks.";
            return;
        }
        if (await SaveAsync(settings.WithUrl(url))) Status = url.Length > 0 ? "Webhook saved." : "Webhook removed.";
    }

    /// <summary>Reference BtnClearChatAlertsWebhook_Click: no URL and off.</summary>
    [RelayCommand]
    private async Task ClearUrlAsync()
    {
        UrlDraft = "";
        if (await SaveAsync(settings with { Url = "", Enabled = false })) Status = "Webhook removed.";
    }

    /// <summary>Reference "DiscordWebhook" alert item: only switches on while signed in.</summary>
    [RelayCommand]
    private Task ToggleEnabledAsync() => !Enabled && !IsSignedIn ? Task.CompletedTask : SaveAsync(settings with { Enabled = !Enabled });

    /// <summary>Select all / Deselect all (reference SetAllAlerts sets the profile flag directly).</summary>
    public Task SetEnabledAsync(bool value) => Enabled == value || !HasProfile ? Task.CompletedTask : SaveAsync(settings with { Enabled = value });

    /// <summary>Reference SendDiscordWebhookAsync: enabled, signed in and a URL set; failures are only logged.</summary>
    public async Task SendAsync(string message)
    {
        if (!settings.Enabled || !IsSignedIn || string.IsNullOrWhiteSpace(settings.Url)) return;
        try { await sender.SendAsync(settings.Url, DiscordWebhookSettings.Content(settings.Mention, profile()?.Name, message), settings.Tts, CancellationToken.None); }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or InvalidOperationException)
        { Log?.Invoke("[Discord] Webhook send failed."); }
    }

    private void Change(DiscordWebhookSettings next) => _ = SaveAsync(next);

    private async Task<bool> SaveAsync(DiscordWebhookSettings next)
    {
        if (profile() is not { } target) { Status = "Select a server first."; return false; }
        var previous = settings;
        ++version;
        Apply(next);
        try
        {
            await store.SetAsync(target.Path, target.ProfileKey, next, CancellationToken.None);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or ProfileChangedException or ArgumentException)
        {
            Apply(previous);
            Status = "The webhook could not be saved to the server profile.";
            return false;
        }
    }

    private void Apply(DiscordWebhookSettings next)
    {
        bool enabledChanged = next.Enabled != settings.Enabled;
        settings = next;
        foreach (string name in new[] { nameof(Settings), nameof(Enabled), nameof(Exclusive), nameof(Tts), nameof(MentionEveryone), nameof(MentionHere), nameof(HasProfile) })
            OnPropertyChanged(name);
        if (enabledChanged) EnabledChanged?.Invoke();
    }
}
