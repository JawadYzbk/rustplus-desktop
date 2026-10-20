using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// The suite runs in English, whatever language the machine it runs on is set to.
///
/// The app reads its strings in the language the user chose and follows the operating system's language when nothing
/// is chosen, so a test that renders a label would otherwise change its answer with the machine's language - and a
/// suite that only passes on an English machine is not a suite that passes.
/// </summary>
[TestClass]
public sealed class TestCulture
{
    /// <summary>
    /// The language the suite reads its strings in. It is the app's own English entry, so the suite reads the same
    /// catalogue the app does when the user picks English.
    /// </summary>
    public static readonly CultureInfo English = new("en-US");

    [AssemblyInitialize]
    public static void UseEnglish(TestContext context)
    {
        CultureInfo.DefaultThreadCurrentUICulture = English;
        CultureInfo.CurrentUICulture = English;
        CultureInfo.DefaultThreadCurrentCulture = English;
        CultureInfo.CurrentCulture = English;
    }
}
