namespace RustPlusDesk.Core.Alarms;

/// <summary>Show: new alarm with a popup; Accepted: new generic alarm without a popup; Update: refines a recent row; Drop: duplicate.</summary>
public enum AlarmDecision { Show, Update, Drop, Accepted }

/// <summary>
/// Reference alarm de-duplication (MainWindow alarm handling): the same entity within 5 s is shown once,
/// though a detailed message replaces the generic "Alarm activated!" row; an alarm without an entity ID
/// arriving within 5 s of any alarm is dropped (or updates it when detailed); generic alarms only pop up
/// when the GenericAlarmPopupEnabled setting (default off) allows it.
/// </summary>
public sealed class AlarmGate
{
    public const string GenericMessage = "Alarm activated!";
    public static readonly TimeSpan Window = TimeSpan.FromSeconds(5);
    private readonly Dictionary<uint, DateTimeOffset> lastByEntity = [];
    private DateTimeOffset lastAny = DateTimeOffset.MinValue;

    public AlarmDecision Decide(uint? entityId, string message, DateTimeOffset now, bool genericPopupEnabled)
    {
        bool detailed = !string.Equals(message?.Trim(), GenericMessage, StringComparison.Ordinal);
        if (entityId is { } id)
        {
            if (lastByEntity.TryGetValue(id, out var last) && now - last < Window)
                return detailed ? AlarmDecision.Update : AlarmDecision.Drop;
            lastByEntity[id] = now;
            lastAny = now;
            return AlarmDecision.Show;
        }
        if (now - lastAny < Window) return detailed ? AlarmDecision.Update : AlarmDecision.Drop;
        lastAny = now;
        // Reference: a generic alarm that passed de-duplication still plays audio and posts chat alerts; only the popup is optional.
        return genericPopupEnabled ? AlarmDecision.Show : AlarmDecision.Accepted;
    }
}
