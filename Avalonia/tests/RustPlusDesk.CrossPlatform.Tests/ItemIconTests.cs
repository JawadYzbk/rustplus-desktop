using Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Shared.Controls;
using RustPlusDesk.Infrastructure.Items;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ItemIconTests
{
    [TestMethod]
    public async Task CustomDeviceIcons_ResolveLegacyIds_ResetToKindDefault_AndDecodeAllAvailableArtwork()
    {
        await using var session = HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder));
        await session.Dispatch(() =>
        {
            var catalog = BundledItemCatalog.Icons;
            Assert.AreEqual(1245, catalog.Count); // The reference data contains two names for ID 1115193056; its last entry wins.
            int missing = 0;
            foreach (var item in catalog.Values)
                if (ItemIcon.Load(item.ShortName) is null) missing++;
            Assert.AreEqual(22, missing, "The reference pack does not contain these canonical item images; preserve choices with a fallback.");
            var control = new ItemIcon { DefaultShortName = "smart.alarm", ItemId = 988652725 };
            Assert.AreSame(ItemIcon.Load("smart.switch"), control.Source);
            control.ItemId = null;
            Assert.AreSame(ItemIcon.Load("smart.alarm"), control.Source);
            control.ShortName = "../unavailable-icon";
            Assert.AreSame(ItemIcon.Load("smart.alarm"), control.Source);
            return true;
        }, CancellationToken.None);
    }

    [TestMethod]
    public async Task All796RequiredInputAndOutputIcons_AreBundledDecodableAndShared()
    {
        await using var session = HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder));
        await session.Dispatch(() =>
        {
            var catalog = new BundledRecyclerCatalog().Load();
            string[] keys = catalog.Items.Select(item => item.Resource.ShortName)
                .Concat(catalog.Items.SelectMany(item => item.Wild.Concat(item.Safe)).Select(yield => yield.ShortName))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            Assert.AreEqual(796, keys.Length);
            foreach (string key in keys)
            {
                var icon = ItemIcon.Load(key);
                Assert.IsNotNull(icon, key);
                Assert.AreEqual(40, icon.PixelSize.Width, key);
                Assert.AreSame(icon, ItemIcon.Load(key), key);
            }
            Assert.IsNull(ItemIcon.Load("unknown-item"));
            return true;
        }, CancellationToken.None);
    }
}
