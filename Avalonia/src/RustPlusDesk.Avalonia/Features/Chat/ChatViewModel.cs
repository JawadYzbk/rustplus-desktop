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
}

/// <summary>
/// Team and clan chat for the connected server (reference chat overlay): history on connect, live
/// broadcasts, Enter to send, the 128-character in-game limit and a "not sent" error that keeps the draft.
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
    private ChatChannel channel = ChatChannel.Team;
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

    public ObservableCollection<ChatLine> Lines => lines[Channel];
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
