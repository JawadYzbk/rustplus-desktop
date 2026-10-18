using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.Avalonia.Features.Chat;

public sealed record ChatLine(ChatMessage Message)
{
    public string Time => Message.Time.ToLocalTime().ToString("HH:mm");
    public string Name => string.IsNullOrWhiteSpace(Message.Name) ? Message.SteamId.ToString(System.Globalization.CultureInfo.InvariantCulture) : Message.Name;
    public string Text => Message.Text;
    public bool HasSteamId => Message.SteamId != 0;
    public string SteamId => Message.SteamId.ToString(System.Globalization.CultureInfo.InvariantCulture);
    public Uri? SteamProfile => HasSteamId ? new Uri("https://steamcommunity.com/profiles/" + SteamId) : null;

    /// <summary>Reference chat search: the trimmed query in the text or the author, ignoring case.</summary>
    public bool Matches(string query) => string.IsNullOrWhiteSpace(query)
        || Text.Contains(query.Trim(), StringComparison.OrdinalIgnoreCase) || Name.Contains(query.Trim(), StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// Team and clan chat for the connected server (reference chat overlay): history on connect, live
/// broadcasts, Enter to send, the 128-character in-game limit and a "not sent" error that keeps the draft.
/// Search filters by text or author; the message menu copies text or Steam ID, mentions the author and
/// centres the map on them.
/// </summary>
public sealed partial class ChatViewModel : ObservableObject
{
    private const int MaxLines = 300;
    private readonly Dictionary<ChatChannel, ObservableCollection<ChatLine>> lines = new()
    {
        [ChatChannel.Team] = [],
        [ChatChannel.Clan] = []
    };
    private IServerChatSession? session;
    private CancellationToken sessionToken;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Lines))]
    [NotifyPropertyChangedFor(nameof(IsClan))]
    [NotifyPropertyChangedFor(nameof(HasNoMatches))]
    private ChatChannel channel = ChatChannel.Team;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Lines))]
    [NotifyPropertyChangedFor(nameof(HasNoMatches))]
    private string searchText = "";
    [ObservableProperty] private bool isSearchOpen;
    /// <summary>Short confirmation or hint from the message menu (copied, not on the map).</summary>
    [ObservableProperty] private string notice = "";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CharacterCount))]
    [NotifyPropertyChangedFor(nameof(IsAtLimit))]
    [NotifyCanExecuteChangedFor(nameof(SendCommand))]
    private string draft = "";
    [ObservableProperty] private string error = "";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SendCommand))]
    private bool isConnected;
    [ObservableProperty] private bool hasClan;
    /// <summary>Chat-alert switch and menu shown in the chat header (set by the shell).</summary>
    [ObservableProperty] private RustPlusDesk.Avalonia.Features.Alerts.ChatAlertsViewModel? alerts;

    // Emoji (reference ChatEmojiInputHelper / ChatEmojiPickerPopup).
    /// <summary>Custom emojis and Rust items usable as ":token:" in chat.</summary>
    [ObservableProperty] private ChatEmojiCatalog emoji = new([]);
    [ObservableProperty] private IReadOnlyList<ChatEmoji> emojiSuggestions = [];
    [ObservableProperty] private ChatEmoji? selectedEmojiSuggestion;
    [ObservableProperty] private bool isEmojiSuggestionsOpen;
    [ObservableProperty] private string emojiPickerQuery = "";
    [ObservableProperty] private IReadOnlyList<ChatEmoji> emojiPickerItems = ChatEmojiCatalog.Custom;
    public bool HasNoEmojiPickerItems => EmojiPickerItems.Count == 0;
    private int emojiStart = -1;

    partial void OnEmojiPickerQueryChanged(string value) => RefreshEmojiPicker();
    partial void OnEmojiPickerItemsChanged(IReadOnlyList<ChatEmoji> value) => OnPropertyChanged(nameof(HasNoEmojiPickerItems));

    /// <summary>Reference picker: the custom emojis by default, the whole catalogue (60) once something is typed.</summary>
    public void RefreshEmojiPicker()
    {
        string query = EmojiPickerQuery.Trim();
        EmojiPickerItems = query.Length == 0 ? ChatEmojiCatalog.Custom : Emoji.Search(query, ChatEmojiCatalog.PickerResults);
    }

    /// <summary>Reference CheckAutocomplete after each edit: up to 15 matches for the ":query" before the caret.</summary>
    public void UpdateEmojiAutocomplete(int caret)
    {
        if (ChatEmojiCatalog.AutocompleteQuery(Draft, caret) is { } found && Emoji.Search(found.Query, ChatEmojiCatalog.AutocompleteResults) is { Count: > 0 } results)
        {
            emojiStart = found.Start;
            EmojiSuggestions = results;
            SelectedEmojiSuggestion = results[0];
            IsEmojiSuggestionsOpen = true;
        }
        else CloseEmojiSuggestions();
    }

    public void CloseEmojiSuggestions()
    {
        IsEmojiSuggestionsOpen = false;
        emojiStart = -1;
    }

    public void MoveEmojiSuggestion(int delta)
    {
        if (EmojiSuggestions.Count == 0) return;
        int index = SelectedEmojiSuggestion is { } current ? EmojiSuggestions.ToList().IndexOf(current) : -1;
        SelectedEmojiSuggestion = EmojiSuggestions[((index + delta) % EmojiSuggestions.Count + EmojiSuggestions.Count) % EmojiSuggestions.Count];
    }

    /// <summary>Reference ApplyAutocomplete; returns the caret after the inserted tag, or null when nothing changed.</summary>
    public int? AcceptEmojiSuggestion(int caret, ChatEmoji? choice = null)
    {
        if ((choice ?? SelectedEmojiSuggestion) is not { } emoji) { CloseEmojiSuggestions(); return null; }
        var (text, next) = ChatEmojiCatalog.ApplyAutocomplete(Draft, caret, emojiStart, emoji);
        CloseEmojiSuggestions();
        return SetDraftWithinLimit(text, next);
    }

    /// <summary>Reference picker selection: the tag and a space at the caret.</summary>
    public int? InsertEmoji(ChatEmoji emoji, int caret)
    {
        var (text, next) = ChatEmojiCatalog.Insert(Draft, caret, emoji);
        return SetDraftWithinLimit(text, next);
    }

    private int? SetDraftWithinLimit(string text, int caret)
    {
        if (text.Length > ChatLimits.MaxLength) { Notice = "That emoji does not fit in the 128-character limit."; return null; }
        Draft = text;
        return caret;
    }
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SendCommand))]
    private bool isSending;

    public IReadOnlyList<ChatLine> Lines => IsFiltering ? lines[Channel].Where(line => line.Matches(SearchText)).ToList() : lines[Channel];
    private bool IsFiltering => !string.IsNullOrWhiteSpace(SearchText);
    public bool HasNoMatches => IsFiltering && Lines.Count == 0;
    /// <summary>Centres the live map on a player's Steam ID; false when they are not on the map.</summary>
    public Func<ulong, bool>? CenterOnMap { get; init; }
    public bool IsClan { get => Channel == ChatChannel.Clan; set => Channel = value ? ChatChannel.Clan : ChatChannel.Team; }
    public string CharacterCount => $"{Draft.Length}/{ChatLimits.MaxLength}";
    public bool IsAtLimit => Draft.Length >= ChatLimits.MaxLength;

    partial void OnDraftChanged(string value)
    {
        if (value.Length > ChatLimits.MaxLength) Draft = value[..ChatLimits.MaxLength];
    }

    public void Reset()
    {
        session = null;
        IsConnected = false;
        HasClan = false;
        Error = "";
        Channel = ChatChannel.Team;
        Notice = "";
        IsSearchOpen = false;
        SearchText = "";
        foreach (var list in lines.Values) list.Clear();
    }

    public async Task RunAsync(IServerChatSession connected, CancellationToken ct)
    {
        session = connected;
        sessionToken = ct;
        IsConnected = true;
        try
        {
            await LoadAsync(connected, ChatChannel.Team, ct);
            // Servers without clans answer the clan request with an error; the clan tab then stays hidden.
            HasClan = await LoadAsync(connected, ChatChannel.Clan, ct);
            await foreach (var message in connected.WatchChatAsync(ct)) Add(message);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
        finally
        {
            if (ReferenceEquals(session, connected)) { session = null; IsConnected = false; }
        }
    }

    private async Task<bool> LoadAsync(IServerChatSession source, ChatChannel target, CancellationToken ct)
    {
        try
        {
            foreach (var message in await source.ReadChatAsync(target, ct).WaitAsync(ct)) Add(message);
            return true;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception) { return false; }
    }

    internal void Add(ChatMessage message)
    {
        if (message.Channel == ChatChannel.Clan) HasClan = true;
        var list = lines[message.Channel];
        // History and the live broadcast can both deliver the same message.
        if (list.Any(line => line.Message.SteamId == message.SteamId && line.Message.Text == message.Text
                && Math.Abs((line.Message.Time - message.Time).TotalSeconds) < 2)) return;
        int index = list.Count;
        while (index > 0 && list[index - 1].Message.Time > message.Time) index--;
        list.Insert(index, new ChatLine(message));
        while (list.Count > MaxLines) list.RemoveAt(0);
        if (IsFiltering && message.Channel == Channel)
        {
            OnPropertyChanged(nameof(Lines));
            OnPropertyChanged(nameof(HasNoMatches));
        }
    }

    /// <summary>Opening focuses the box in the view; closing clears the query like the reference toggle.</summary>
    [RelayCommand]
    private void ToggleSearch()
    {
        IsSearchOpen = !IsSearchOpen;
        if (!IsSearchOpen) SearchText = "";
    }

    [RelayCommand] private void ClearSearch() => SearchText = "";

    /// <summary>Prefixes the draft with "@author " (reference Mention @User); the in-game limit still applies.</summary>
    public void Mention(ChatLine line)
    {
        if (string.IsNullOrWhiteSpace(line.Message.Name)) return;
        Draft = $"@{line.Message.Name} " + Draft;
    }

    [RelayCommand]
    private void CenterOnAuthor(ChatLine? line)
    {
        if (line is not { HasSteamId: true }) return;
        Notice = CenterOnMap?.Invoke(line.Message.SteamId) == true ? "" : "Teammate is not currently visible on the active map.";
    }

    /// <summary>
    /// Posts an automated alert (shop alerts) to team chat, cut to the in-game limit. Returns false while disconnected.
    /// Pending: the reference's clan-channel choice, automated-chat master switch and Discord forwarding.
    /// </summary>
    public async Task<bool> SendAutomatedAsync(string text)
    {
        if (session is not { } target || string.IsNullOrWhiteSpace(text)) return false;
        string message = text.Length > ChatLimits.MaxLength ? text[..ChatLimits.MaxLength] : text;
        await target.SendChatAsync(ChatChannel.Team, message, sessionToken);
        return true;
    }

    private bool CanSend => IsConnected && !IsSending && !string.IsNullOrWhiteSpace(Draft);

    [RelayCommand(CanExecute = nameof(CanSend))]
    private async Task SendAsync()
    {
        if (session is not { } target) return;
        string text = Draft.Trim();
        IsSending = true;
        Error = "";
        try
        {
            await target.SendChatAsync(Channel, text, sessionToken);
            Draft = "";
        }
        catch (OperationCanceledException) { Error = "Message not sent: the connection closed."; }
        catch (Exception) { Error = "Message not sent. Check the connection and try again."; }
        finally { IsSending = false; }
    }
}
