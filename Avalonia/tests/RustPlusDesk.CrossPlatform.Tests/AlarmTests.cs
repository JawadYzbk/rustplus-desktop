using Avalonia.Controls;
using Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alarms;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Alarms;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class AlarmTests
{
    private static readonly DateTimeOffset T0 = new(2026, 10, 18, 12, 0, 0, TimeSpan.Zero);

    private static CompanionUpdate Alarm(uint? entity, string message, double seconds, string server = "EU Vanilla")
        => new(CompanionUpdateKind.Alarm, "Base alarm", message) { EntityId = entity, ServerName = server, OccurredAt = T0.AddSeconds(seconds) };

    [TestMethod]
    public void Gate_FollowsReferenceFiveSecondRules()
    {
        var gate = new AlarmGate();
        Assert.AreEqual(AlarmDecision.Show, gate.Decide(7, AlarmGate.GenericMessage, T0, false));
        Assert.AreEqual(AlarmDecision.Drop, gate.Decide(7, AlarmGate.GenericMessage, T0.AddSeconds(2), false));
        Assert.AreEqual(AlarmDecision.Update, gate.Decide(7, "Raid on main base!", T0.AddSeconds(3), false));
        Assert.AreEqual(AlarmDecision.Show, gate.Decide(7, AlarmGate.GenericMessage, T0.AddSeconds(5), false), "The window has passed.");
        // Generic (no entity) alarms: dropped near another alarm, otherwise only shown when enabled.
        Assert.AreEqual(AlarmDecision.Drop, gate.Decide(null, AlarmGate.GenericMessage, T0.AddSeconds(6), true));
        Assert.AreEqual(AlarmDecision.Drop, gate.Decide(null, AlarmGate.GenericMessage, T0.AddSeconds(20), false));
        Assert.AreEqual(AlarmDecision.Show, gate.Decide(null, AlarmGate.GenericMessage, T0.AddSeconds(40), true));
        Assert.AreEqual(AlarmDecision.Show, gate.Decide(8, AlarmGate.GenericMessage, T0.AddSeconds(41), false), "Another entity is separate.");
    }

    [TestMethod]
    public void ViewModel_ShowsNewestFirst_RefinesGenericRows_AndIgnoresOtherKinds()
    {
        var model = new AlarmsViewModel();
        int opened = 0;
        model.OpenRequested += (_, _) => opened++;
        model.Handle(new CompanionUpdate(CompanionUpdateKind.Pairing, "Paired", "Server added"));
        Assert.AreEqual(0, model.Items.Count);
        model.Handle(Alarm(7, AlarmGate.GenericMessage, 0));
        model.Handle(Alarm(7, "Raid on main base!", 2));
        Assert.AreEqual(1, model.Items.Count);
        Assert.AreEqual("Raid on main base!", model.Items[0].Message);
        model.Handle(Alarm(9, "Turret base", 10));
        Assert.AreEqual("Turret base", model.Items[0].Message, "Newest first.");
        Assert.AreEqual(2, opened);
        Assert.IsTrue(model.IsOpen);
        model.ClearCommand.Execute(null);
        Assert.AreEqual(0, model.Items.Count);
        model.CloseCommand.Execute(null);
        Assert.IsFalse(model.IsOpen);
    }

    [TestMethod]
    public async Task Window_RendersAlarms_AndClosesFromTheViewModel()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new AlarmsViewModel();
            model.Handle(Alarm(7, "Raid on main base!", 0));
            var window = new AlarmWindow { DataContext = model };
            window.Show();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame!.Save(Path.Combine(AppContext.BaseDirectory, "ui-alarms.png"), global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
            bool closed = false;
            window.Closed += (_, _) => closed = true;
            model.CloseCommand.Execute(null);
            Assert.IsTrue(closed);
            await Task.CompletedTask;
            return true;
        }, default);
    }
}
