using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Core.Devices;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeviceFilterTests
{
    private static readonly PairedDevice Lights = new(11, "Compound lights", DeviceKind.Switch) { Alias = "Lights" };
    private static readonly PairedDevice Alarm = new(22, "Base alarm", DeviceKind.Alarm);
    private static readonly PairedDevice Tc = new(33, "Tool cupboard", DeviceKind.StorageMonitor);
    private static readonly PairedDevice Group = new(0x80000001, "Main base", DeviceKind.Group) { Children = [Lights, Alarm] };

    [TestMethod]
    [DataRow("", DeviceKindFilter.All, true)]
    [DataRow("lights", DeviceKindFilter.All, true)]
    [DataRow("LIGHTS", DeviceKindFilter.Switch, true)]
    [DataRow("lights", DeviceKindFilter.Alarm, false)]
    [DataRow("switch", DeviceKindFilter.All, true)]
    [DataRow("11", DeviceKindFilter.All, true)]
    [DataRow("cupboard", DeviceKindFilter.All, false)]
    public void Matches_FollowsReferenceSearchAndKindRules(string search, DeviceKindFilter kind, bool expected)
        => Assert.AreEqual(expected, DeviceFilter.Matches(Lights, search, kind));

    [TestMethod]
    public void Groups_StayVisibleAroundMatchingChildren_AndHideOtherwise()
    {
        Assert.IsTrue(DeviceFilter.Matches(Group, "alarm", DeviceKindFilter.All));
        Assert.IsTrue(DeviceFilter.Matches(Group, "", DeviceKindFilter.Alarm));
        Assert.IsFalse(DeviceFilter.Matches(Group, "", DeviceKindFilter.Storage));
        Assert.IsTrue(DeviceFilter.Matches(Group, "main", DeviceKindFilter.All));
    }

    [TestMethod]
    public void DevicesViewModel_FlagsRowsAndReportsNoMatches()
    {
        var model = new DevicesViewModel();
        model.Select([Group, Tc]);
        var rows = model.Devices.SelectMany(row => row.DescendantsAndSelf()).ToDictionary(row => row.Definition.EntityId);
        Assert.IsTrue(rows.Values.All(row => row.IsMatch));
        model.KindFilter = DeviceKindFilter.Storage;
        Assert.IsFalse(rows[Group.EntityId].IsMatch);
        Assert.IsTrue(rows[Tc.EntityId].IsMatch);
        model.KindFilter = DeviceKindFilter.All;
        model.SearchText = "base alarm";
        Assert.IsTrue(rows[Group.EntityId].IsMatch);
        Assert.IsTrue(rows[Alarm.EntityId].IsMatch);
        Assert.IsFalse(rows[Lights.EntityId].IsMatch);
        Assert.IsFalse(model.HasNoMatches);
        model.SearchText = "nothing like this";
        Assert.IsTrue(model.HasNoMatches);
        model.Select([Tc]);
        Assert.IsTrue(model.HasNoMatches, "A new server keeps the current filter and re-evaluates its rows.");
        model.SearchText = "";
        Assert.IsFalse(model.HasNoMatches);
    }
}
