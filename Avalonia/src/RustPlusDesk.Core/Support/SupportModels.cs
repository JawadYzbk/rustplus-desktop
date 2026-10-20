namespace RustPlusDesk.Core.Support;

/// <summary>One row in the user's ticket list (reference TicketSummary).</summary>
public sealed record TicketSummary(
    string Id,
    string Category,
    string Status,
    string? Resolution,
    string Subject,
    bool HasUnread,
    DateTimeOffset? LastActivityAt);

/// <summary>A file hung off a ticket or one of its replies (reference TicketAttachment).</summary>
public sealed record TicketAttachment(string Id, string Name, long Size, string Mime, string? Url = null, string? ThumbUrl = null);

/// <summary>One line in a ticket's thread (reference TicketMessage).</summary>
public sealed record TicketMessage(
    string Id,
    string Kind,
    string Body,
    string? AuthorName,
    bool IsStaff,
    IReadOnlyList<TicketAttachment> Attachments,
    DateTimeOffset? CreatedAt);

/// <summary>The full ticket, thread included (reference TicketDetail).</summary>
public sealed record TicketDetail(
    string Id,
    string Category,
    string Status,
    string? Resolution,
    string Subject,
    string Body,
    IReadOnlyList<TicketAttachment> Attachments,
    IReadOnlyList<TicketMessage> Messages,
    DateTimeOffset? CreatedAt);

/// <summary>An account's active mute, offered as something to appeal (reference AppealableSanction).</summary>
public sealed record AppealableSanction(string Id, string Kind, string Reason, DateTimeOffset? ExpiresAt);

/// <summary>What the new-ticket form needs: the categories and the current mute, if any (reference TicketMeta).</summary>
public sealed record TicketMeta(IReadOnlyList<string> Categories, AppealableSanction? Appeal);
