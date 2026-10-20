using System.Globalization;
using Avalonia.Headless;
using Avalonia.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>Keeps framework-cached decorations and every rendered control on the same UI thread.</summary>
[TestClass]
public sealed class UiTestSession
{
    private static readonly Lazy<HeadlessUnitTestSession> Shared = new(() => HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder)));
    internal static UiSession Current => new(Shared.Value);

    [AssemblyCleanup]
    public static async Task CleanupAsync()
    {
        if (Shared.IsValueCreated) await Shared.Value.DisposeAsync();
    }
}

/// <summary>
/// The user interface session the tests draw in.
///
/// It runs the body on the user interface thread and checks that the suite is still in the language it runs in: the
/// app reads its strings in the language the user chose, and a test that changes that language and does not put it
/// back would otherwise change the tests that run after it - quietly, and only on the machine that ran it.
/// </summary>
internal sealed class UiSession(HeadlessUnitTestSession session)
{
    public Task Dispatch(Action action, CancellationToken cancellationToken = default) =>
        session.Dispatch(() => { Check(); action(); }, cancellationToken);

    public Task<T> Dispatch<T>(Func<T> action, CancellationToken cancellationToken = default) =>
        session.Dispatch(() => { Check(); return action(); }, cancellationToken);

    public Task<T> Dispatch<T>(Func<Task<T>> action, CancellationToken cancellationToken = default) =>
        session.Dispatch(async () => { Check(); return await action(); }, cancellationToken);

    private static void Check()
    {
        Assert.AreEqual(TestCulture.English.Name, CultureInfo.CurrentUICulture.Name,
            "The suite runs in English; a test that changes the language puts it back.");
        Assert.AreEqual(TestCulture.English.Name, RustPlusDesk.Core.Localization.Loc.Culture?.Name,
            "The app reads its strings in the suite's language; a test that changes the language puts it back.");
    }
}
