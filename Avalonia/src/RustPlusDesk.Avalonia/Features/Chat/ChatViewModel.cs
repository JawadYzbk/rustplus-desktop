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
