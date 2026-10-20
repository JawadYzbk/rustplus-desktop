using System.Net;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Support;
using RustPlusDesk.Core.Support;
using RustPlusDesk.Infrastructure.Support;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference SupportApi: the user's tickets are read from the cloud API, a ticket is filed with its attachments
/// and a reply is added to the thread.
/// </summary>
[TestClass]
public sealed class SupportTests
{
    private sealed class StubHandler(Func<HttpRequestMessage, (HttpStatusCode Status, string Body)> answer) : HttpMessageHandler
    {
        public List<HttpRequestMessage> Requests { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            var (status, body) = answer(request);
            return new HttpResponseMessage(status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };
        }
    }

    private static SupportService Service(HttpMessageHandler handler, out StubHandler stub)
    {
        stub = new StubHandler(_ => (HttpStatusCode.OK, "{}"));
        return new SupportService(new HttpClient(handler), Path.Combine(Path.GetTempPath(), "rpd-support-none.json"), "0.1.0");
    }

    [TestMethod]
    public async Task TheTicketsAreRead()
    {
        var service = new SupportService(new HttpClient(new StubHandler(_ => (HttpStatusCode.OK, """
            {"data":[{"id":"t1","category":"bug","status":"open","subject":"The map is blank",
              "has_unread":true,"last_activity_at":"2026-10-19T10:00:00+00:00"}]}
            """))), Path.Combine(Path.GetTempPath(), "rpd-support-none.json"), "0.1.0");

        var tickets = await service.GetTicketsAsync(default);
        Assert.AreEqual(1, tickets.Count);
        Assert.AreEqual("t1", tickets[0].Id);
        Assert.AreEqual("The map is blank", tickets[0].Subject);
        Assert.AreEqual("open", tickets[0].Status);
        Assert.IsTrue(tickets[0].HasUnread);
        Assert.IsNotNull(tickets[0].LastActivityAt);
    }

    [TestMethod]
    public async Task ATicketIsReadWithItsThread()
    {
        var service = new SupportService(new HttpClient(new StubHandler(_ => (HttpStatusCode.OK, """
            {"data":{"id":"t1","category":"bug","status":"open","subject":"The map is blank","body":"It is blank.",
              "messages":[{"id":"m1","kind":"reply","body":"Looking into it","author_name":"Support","is_staff":true,
                "created_at":"2026-10-19T11:00:00+00:00"}],
              "attachments":[{"id":"a1","name":"shot.png","size":12,"mime":"image/png"}]}}
            """))), Path.Combine(Path.GetTempPath(), "rpd-support-none.json"), "0.1.0");

        var ticket = await service.GetTicketAsync("t1", default);
        Assert.IsNotNull(ticket);
        Assert.AreEqual("It is blank.", ticket!.Body);
        Assert.AreEqual(1, ticket.Messages.Count);
        Assert.AreEqual("Support", ticket.Messages[0].AuthorName);
        Assert.IsTrue(ticket.Messages[0].IsStaff);
        Assert.AreEqual(1, ticket.Attachments.Count);
        Assert.AreEqual("shot.png", ticket.Attachments[0].Name);
    }

    [TestMethod]
    public async Task TheCategoriesAndTheMuteAreRead()
    {
        var service = new SupportService(new HttpClient(new StubHandler(_ => (HttpStatusCode.OK, """
            {"categories":["bug","question","other"],
             "appealable_sanction":{"id":"s1","kind":"timeout","reason":"too many requests","expires_at":"2026-10-25T10:00:00+00:00"}}
            """))), Path.Combine(Path.GetTempPath(), "rpd-support-none.json"), "0.1.0");

        var meta = await service.GetMetaAsync(default);
        CollectionAssert.AreEqual(new[] { "bug", "question", "other" }, meta.Categories.ToArray());
        Assert.IsNotNull(meta.Appeal);
        Assert.AreEqual("s1", meta.Appeal!.Id);
        Assert.AreEqual("timeout", meta.Appeal.Kind);
    }

    [TestMethod]
    public async Task ATicketIsFiled()
    {
        var service = new SupportService(new HttpClient(new StubHandler(_ => (HttpStatusCode.Created, "{}"))),
            Path.Combine(Path.GetTempPath(), "rpd-support-none.json"), "0.1.0");
        Assert.IsTrue(await service.CreateTicketAsync("bug", "The map is blank", "It is blank since yesterday"));
    }

    [TestMethod]
    public async Task AReplyIsSent()
    {
        var service = new SupportService(new HttpClient(new StubHandler(_ => (HttpStatusCode.OK, "{}"))),
            Path.Combine(Path.GetTempPath(), "rpd-support-none.json"), "0.1.0");
        Assert.IsTrue(await service.ReplyAsync("t1", "Any news?"));
    }

    [TestMethod]
    public async Task TheCard_ReadsTheTickets_AndFilesANewOne()
    {
        var model = new SupportViewModel(
            _ => Task.FromResult<IReadOnlyList<TicketSummary>>(
                [new TicketSummary("t1", "bug", "open", null, "The map is blank", true, null)]),
            _ => Task.FromResult(new TicketMeta(["bug"], null)),
            (category, subject, body, _) => Task.FromResult(true),
            (ticket, answer, _) => Task.FromResult(true),
            (ticket, _) => Task.FromResult<TicketDetail?>(new TicketDetail("t1", "bug", "open", null, "The map is blank", "It is blank", [],
                [new TicketMessage("m1", "reply", "Looking into it", "Support", true, [], null)], null)));

        await model.LoadCommand.ExecuteAsync(null);
        Assert.AreEqual(1, model.Tickets.Count);
        Assert.AreEqual("The map is blank", model.Tickets[0].Subject);
        Assert.AreEqual("bug", model.Category, "The first category is chosen.");

        model.Subject = "A new problem";
        model.Body = "It started today";
        await model.SendCommand.ExecuteAsync(null);
        Assert.AreEqual("The ticket was sent.", model.Status);
        Assert.AreEqual("", model.Subject, "The form is cleared after the ticket is sent.");
    }

    [TestMethod]
    public async Task TheCard_RefusesATicketWithoutASubjectOrAMessage()
    {
        var model = new SupportViewModel(null, null, null, null, null);
        model.Subject = "";
        model.Body = "";
        await model.SendCommand.ExecuteAsync(null);
        StringAssert.Contains(model.Status, "needed");
    }
}
