using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Activity;
using RustPlusDesk.Core.Diagnostics;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ActivityLogTests
{
    private sealed class FixedClock(DateTimeOffset now) : TimeProvider
    {
        public DateTimeOffset Now { get; set; } = now;
        public override DateTimeOffset GetUtcNow() => Now;
        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }

    [TestMethod]
    public void Log_TimestampsTrimsAndCapsAtTheReferenceLimit()
    {
        var clock = new FixedClock(new DateTimeOffset(2026, 10, 18, 9, 5, 7, 42, TimeSpan.Zero));
        var log = new ActivityLog(clock);
        Assert.IsFalse(log.Append("   "));
        Assert.IsFalse(log.Append(null));
        Assert.IsTrue(log.Append("  [listener] started  "));
        Assert.AreEqual("[09:05:07.042] [listener] started", log.Lines.Single());
        for (int i = 0; i < ActivityLog.MaxLines + 5; i++) log.Append("line " + i);
        Assert.AreEqual(ActivityLog.MaxLines, log.Lines.Count);
        StringAssert.EndsWith(log.Lines[0], "line 5", "The oldest lines are dropped first.");
        StringAssert.EndsWith(log.Lines[^1], "line " + (ActivityLog.MaxLines + 4));
        Assert.AreEqual(1, log.Filter(" LINE 1999 ").Count(), "The filter is trimmed and ignores case.");
        Assert.AreEqual(6, log.Filter("line 200").Count(), "line 200 and line 2000-2004.");
        Assert.AreEqual(ActivityLog.MaxLines, log.Filter("").Count());
    }

    [TestMethod]
    [DataRow("playerToken=-123456789 ok", "playerToken=*** ok")]
    [DataRow("\"refresh_token\": \"abc.def\"", "\"refresh_token\": \"***\"")]
    [DataRow("password: hunter2", "password: ***")]
    [DataRow("Authorization: Bearer abc123.xyz", "Authorization: Bearer ***")]
    [DataRow("jwt eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0In0.c2lnbmF0dXJl end", "jwt *** end")]
    [DataRow("api_key=K1&x=2", "api_key=***&x=2")]
    [DataRow("Connected to 203.0.113.5:28082", "Connected to 203.0.113.5:28082")]
    public void Redactor_MasksCredentialsOnly(string input, string expected) => Assert.AreEqual(expected, LogRedactor.Redact(input));

    [TestMethod]
    public void ViewModel_FiltersExpandsAndReportsEmptyStates()
    {
        var model = new ActivityLogViewModel(new FixedClock(new DateTimeOffset(2026, 10, 18, 12, 0, 0, TimeSpan.Zero)));
        Assert.IsTrue(model.IsEmpty);
        Assert.AreEqual("", model.Text);
        var changed = new List<string?>();
        model.PropertyChanged += (_, e) => changed.Add(e.PropertyName);
        model.Append("[connection] EU: Connected");
        model.Append("[tools] secret=x");
        CollectionAssert.Contains(changed, nameof(ActivityLogViewModel.Text));
        Assert.IsFalse(model.IsEmpty);
        Assert.AreEqual("[12:00:00.000] [connection] EU: Connected" + Environment.NewLine + "[12:00:00.000] [tools] secret=***", model.Text);
        model.Filter = "TOOLS";
        Assert.AreEqual("[12:00:00.000] [tools] secret=***", model.Text);
        model.Filter = "nothing";
        Assert.IsTrue(model.HasNoMatches);
        Assert.AreEqual("", model.Text);
        model.ClearFilterCommand.Execute(null);
        Assert.IsFalse(model.HasNoMatches);
        Assert.AreEqual("Expand", model.ExpandCaption);
        model.ToggleExpandCommand.Execute(null);
        Assert.IsTrue(model.IsExpanded);
        Assert.AreEqual("Collapse", model.ExpandCaption);
    }
}
