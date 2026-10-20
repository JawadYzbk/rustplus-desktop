using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Support;

namespace RustPlusDesk.Infrastructure.Support;

/// <summary>
/// Reference SupportApi: the user's support tickets, read and written through the cloud API the account already
/// uses. A ticket can carry attachments, so the new-ticket form is a multipart post.
/// </summary>
public sealed class SupportService(HttpClient http, string cachePath, string clientVersion)
{
    private const string Base = "https://rustplusdesktop.cloud/api/v1/";

    /// <summary>The categories the form offers and the current mute, if there is one.</summary>
    public async Task<TicketMeta> GetMetaAsync(CancellationToken ct)
    {
        try
        {
            var root = await GetAsync("tickets/meta", ct).ConfigureAwait(false);
            var categories = root["categories"] is JsonArray rows
                ? rows.OfType<JsonValue>()
                    .Select(value => value.TryGetValue(out string? text) ? text : "")
                    .Where(text => text.Length > 0).ToArray()
                : Array.Empty<string>();
            AppealableSanction? appeal = null;
            if (root["appealable_sanction"] is JsonObject sanction)
                appeal = new AppealableSanction(Text(sanction, "id") ?? "", Text(sanction, "kind") ?? "timeout",
                    Text(sanction, "reason") ?? "", Date(sanction, "expires_at"));
            return new TicketMeta(categories, appeal);
        }
        catch (Exception ex) when (ex is HttpRequestException or JsonException or TaskCanceledException)
        {
            return new TicketMeta([], null);
        }
    }

    /// <summary>The user's tickets, newest first.</summary>
    public async Task<IReadOnlyList<TicketSummary>> GetTicketsAsync(CancellationToken ct)
    {
        try
        {
            var root = await GetAsync("tickets", ct).ConfigureAwait(false);
            if (root["data"] is not JsonArray rows) return [];
            return rows.OfType<JsonObject>().Select(ParseSummary).ToArray();
        }
        catch (Exception ex) when (ex is HttpRequestException or JsonException or TaskCanceledException)
        {
            return [];
        }
    }

    /// <summary>One ticket with its thread.</summary>
    public async Task<TicketDetail?> GetTicketAsync(string id, CancellationToken ct)
    {
        try
        {
            var root = await GetAsync($"tickets/{id}", ct).ConfigureAwait(false);
            return root["data"] is JsonObject data ? ParseDetail(data) : null;
        }
        catch (Exception ex) when (ex is HttpRequestException or JsonException or TaskCanceledException)
        {
            return null;
        }
    }

    /// <summary>
    /// Files a ticket with any attachments. It is sent as multipart so a bug report can carry its screenshots
    /// alongside the text.
    /// </summary>
    public async Task<bool> CreateTicketAsync(string category, string subject, string body,
        IReadOnlyDictionary<string, string>? meta = null, IReadOnlyList<string>? filePaths = null,
        string? sanctionId = null, CancellationToken ct = default)
    {
        using var content = new MultipartFormDataContent
        {
            { new StringContent(category), "category" },
            { new StringContent(subject), "subject" },
            { new StringContent(body), "body" }
        };
        if (!string.IsNullOrEmpty(sanctionId)) content.Add(new StringContent(sanctionId), "sanction_id");
        if (meta is not null)
            foreach (var pair in meta) content.Add(new StringContent(pair.Value), $"meta[{pair.Key}]");
        AddFiles(content, filePaths);
        return await PostMultipartAsync("tickets", content, ct).ConfigureAwait(false);
    }

    /// <summary>Adds a reply to a ticket, with any attachments.</summary>
    public async Task<bool> ReplyAsync(string ticketId, string body, IReadOnlyList<string>? filePaths = null,
        CancellationToken ct = default)
    {
        using var content = new MultipartFormDataContent { { new StringContent(body), "body" } };
        AddFiles(content, filePaths);
        return await PostMultipartAsync($"tickets/{ticketId}/messages", content, ct).ConfigureAwait(false);
    }

    private async Task<JsonObject> GetAsync(string route, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, Base + route);
        AddHeaders(request);
        string? token = await TokenAsync(ct).ConfigureAwait(false);
        if (token is not null) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        using var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        await response.Content.LoadIntoBufferAsync(1_048_576, ct).ConfigureAwait(false);
        return JsonNode.Parse(await response.Content.ReadAsByteArrayAsync(ct).ConfigureAwait(false))?.AsObject()
            ?? throw new JsonException();
    }

    private async Task<bool> PostMultipartAsync(string route, MultipartFormDataContent content, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, Base + route) { Content = content };
        AddHeaders(request);
        string? token = await TokenAsync(ct).ConfigureAwait(false);
        if (token is not null) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        using var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
        return response.IsSuccessStatusCode;
    }

    private void AddHeaders(HttpRequestMessage request)
    {
        request.Headers.Add("X-Client-Version", clientVersion);
        request.Headers.UserAgent.Add(new ProductInfoHeaderValue("RustPlusDesktop-Avalonia", clientVersion));
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    /// <summary>The cloud session's token; the tickets belong to the signed-in account.</summary>
    private async Task<string?> TokenAsync(CancellationToken ct)
    {
        try
        {
            var session = await Cloud.CloudSessionFile.ReadAsync(cachePath, ct).ConfigureAwait(false);
            return session.Token.Length == 0 ? null : session.Token;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            return null;
        }
    }

    private static void AddFiles(MultipartFormDataContent content, IReadOnlyList<string>? filePaths)
    {
        if (filePaths is null) return;
        foreach (string path in filePaths)
        {
            if (!File.Exists(path)) continue;
            var bytes = File.ReadAllBytes(path);
            var file = new ByteArrayContent(bytes);
            file.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            content.Add(file, "files", Path.GetFileName(path));
        }
    }

    private static TicketSummary ParseSummary(JsonObject row) => new(
        Text(row, "id") ?? "",
        Text(row, "category") ?? "",
        Text(row, "status") ?? "",
        Text(row, "resolution"),
        Text(row, "subject") ?? "",
        row["has_unread"]?.GetValue<bool>() ?? false,
        Date(row, "last_activity_at"));

    private static TicketDetail ParseDetail(JsonObject row) => new(
        Text(row, "id") ?? "",
        Text(row, "category") ?? "",
        Text(row, "status") ?? "",
        Text(row, "resolution"),
        Text(row, "subject") ?? "",
        Text(row, "body") ?? "",
        Attachments(row["attachments"] as JsonArray),
        Messages(row["messages"] as JsonArray),
        Date(row, "created_at"));

    private static IReadOnlyList<TicketMessage> Messages(JsonArray? rows)
    {
        if (rows is null) return [];
        return rows.OfType<JsonObject>().Select(row => new TicketMessage(
            Text(row, "id") ?? "",
            Text(row, "kind") ?? "",
            Text(row, "body") ?? "",
            Text(row, "author_name"),
            row["is_staff"]?.GetValue<bool>() ?? false,
            Attachments(row["attachments"] as JsonArray),
            Date(row, "created_at"))).ToArray();
    }

    private static IReadOnlyList<TicketAttachment> Attachments(JsonArray? rows)
    {
        if (rows is null) return [];
        return rows.OfType<JsonObject>().Select(row => new TicketAttachment(
            Text(row, "id") ?? "",
            Text(row, "name") ?? "",
            row["size"]?.GetValue<long>() ?? 0,
            Text(row, "mime") ?? "application/octet-stream",
            Text(row, "url"),
            Text(row, "thumb_url"))).ToArray();
    }

    private static string? Text(JsonObject row, string name) =>
        row[name] is JsonValue value && value.TryGetValue(out string? text) ? text : null;

    private static DateTimeOffset? Date(JsonObject row, string name) =>
        row[name] is JsonValue value && value.TryGetValue(out DateTimeOffset when) ? when : null;
}
