using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Support;

namespace RustPlusDesk.Avalonia.Features.Support;

/// <summary>One ticket in the support card's list.</summary>
public sealed record TicketRow(TicketSummary Summary)
{
    public string Id => Summary.Id;
    public string Subject => Summary.Subject.Length > 0 ? Summary.Subject : "(no subject)";
    public string Category => Summary.Category;
    public string Status => Summary.Status;
    public string LastActivity => Summary.LastActivityAt?.ToLocalTime().ToString("g", System.Globalization.CultureInfo.CurrentCulture) ?? "";
    public bool HasUnread => Summary.HasUnread;
}

/// <summary>
/// Reference the support card: the user's tickets, the form that files a new one and the thread of an open ticket.
/// A ticket needs the signed-in cloud account, so without one the card says so instead of pretending.
/// </summary>
public sealed partial class SupportViewModel(
    Func<CancellationToken, Task<IReadOnlyList<TicketSummary>>>? loadTickets,
    Func<CancellationToken, Task<TicketMeta>>? loadMeta,
    Func<string, string, string, CancellationToken, Task<bool>>? createTicket,
    Func<string, string, CancellationToken, Task<bool>>? reply,
    Func<string, CancellationToken, Task<TicketDetail?>>? openTicket) : ObservableObject
{
    [ObservableProperty] private bool isBusy;
    [ObservableProperty] private string status = "";
    [ObservableProperty] private string category = "";
    [ObservableProperty] private string subject = "";
    [ObservableProperty] private string body = "";
    [ObservableProperty] private bool isSending;
    [ObservableProperty] private TicketDetail? open;

    public ObservableCollection<TicketRow> Tickets { get; } = [];
    public ObservableCollection<TicketMessage> Messages { get; } = [];
    public IReadOnlyList<string> Categories { get; private set; } = [];

    /// <summary>Reference Load: the tickets and the categories are read.</summary>
    [RelayCommand]
    public async Task LoadAsync(CancellationToken ct)
    {
        if (loadTickets is null) { Status = "The support service is not available."; return; }
        IsBusy = true;
        try
        {
            var tickets = await loadTickets(ct);
            Tickets.Clear();
            foreach (var ticket in tickets) Tickets.Add(new TicketRow(ticket));
            Status = Tickets.Count == 0 ? "No tickets yet." : $"{Tickets.Count} ticket(s).";

            if (loadMeta is { } meta)
            {
                var information = await meta(ct);
                Categories = information.Categories;
                if (Categories.Count > 0 && Category.Length == 0) Category = Categories[0];
            }
        }
        catch (Exception ex) { Status = $"The tickets could not be read: {ex.Message}"; }
        finally { IsBusy = false; }
    }

    /// <summary>Reference the new-ticket form: the ticket is filed and the list is read again.</summary>
    [RelayCommand]
    public async Task SendAsync(CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(Subject) || string.IsNullOrWhiteSpace(Body))
        {
            Status = "A subject and a message are needed.";
            return;
        }
        if (createTicket is null) { Status = "The support service is not available."; return; }
        IsSending = true;
        try
        {
            bool sent = await createTicket(Category, Subject.Trim(), Body.Trim(), ct);
            if (!sent)
            {
                Status = "The ticket could not be sent.";
                return;
            }
            Subject = "";
            Body = "";
            await LoadAsync(ct);
            Status = "The ticket was sent.";
        }
        catch (Exception ex) { Status = $"The ticket could not be sent: {ex.Message}"; }
        finally { IsSending = false; }
    }

    /// <summary>Reference the ticket thread: the ticket is opened and its replies are shown.</summary>
    [RelayCommand]
    public async Task OpenAsync(TicketRow? row, CancellationToken ct)
    {
        if (row is null || openTicket is null) return;
        IsBusy = true;
        try
        {
            Open = await openTicket(row.Id, ct);
            Messages.Clear();
            if (Open is not null)
                foreach (var message in Open.Messages) Messages.Add(message);
            Status = Open is null ? "The ticket could not be opened." : $"Opened {Open.Subject}.";
        }
        catch (Exception ex) { Status = $"The ticket could not be opened: {ex.Message}"; }
        finally { IsBusy = false; }
    }

    /// <summary>Reference the reply box: the answer is added to the thread.</summary>
    [RelayCommand]
    public async Task ReplyAsync(CancellationToken ct)
    {
        if (reply is null || Open is null) return;
        if (string.IsNullOrWhiteSpace(Body)) return;
        IsSending = true;
        try
        {
            bool sent = await reply(Open.Id, Body.Trim(), ct);
            Status = sent ? "The reply was sent." : "The reply could not be sent.";
            if (sent)
            {
                Body = "";
                await ReloadOpenAsync(ct);
            }
        }
        catch (Exception ex) { Status = $"The reply could not be sent: {ex.Message}"; }
        finally { IsSending = false; }
    }

    /// <summary>Reads the open ticket again, so the new reply is in the thread.</summary>
    private async Task ReloadOpenAsync(CancellationToken ct)
    {
        if (Open is null || openTicket is null) return;
        Open = await openTicket(Open.Id, ct);
        Messages.Clear();
        if (Open is not null)
            foreach (var message in Open.Messages) Messages.Add(message);
    }
}
