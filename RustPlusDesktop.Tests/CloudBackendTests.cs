using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Services.Cloud;

namespace RustPlusDesktop.Tests;

[TestClass]
public class CloudBackendTests
{
    [TestMethod]
    public void Mode_DefaultsToSupabase_ForZeroBehaviourChange()
    {
        Assert.AreEqual(CloudBackendMode.Supabase, CloudBackend.Mode);
        Assert.IsFalse(CloudBackend.UseLaravel);
    }

    [DataTestMethod]
    [DataRow("user-profile/limits", "me/limits")]
    [DataRow("user-profile/presence", "profile/presence")]
    [DataRow("user-profile/consent", "profile/consent")]
    [DataRow("user-profile", "profile")]
    [DataRow("discord-roles", "me/discord/sync-roles")]
    [DataRow("/user-profile/limits/", "me/limits")]
    public void MapEdgeFunctionToRoute_MapsKnownFunctions(string edge, string expected)
    {
        Assert.AreEqual(expected, CloudBackend.MapEdgeFunctionToRoute(edge));
    }

    [DataTestMethod]
    [DataRow("user-profile/claim")]
    [DataRow("unknown-function")]
    [DataRow("")]
    [DataRow("   ")]
    public void MapEdgeFunctionToRoute_ReturnsNull_ForUnmapped(string edge)
    {
        Assert.IsNull(CloudBackend.MapEdgeFunctionToRoute(edge));
    }

    [TestMethod]
    public void ApiUrl_NormalisesSlashes()
    {
        Assert.AreEqual(
            "https://api.example.com/api/v1/profile/presence",
            CloudBackend.ApiUrl("https://api.example.com", "/profile/presence"));
    }
}
