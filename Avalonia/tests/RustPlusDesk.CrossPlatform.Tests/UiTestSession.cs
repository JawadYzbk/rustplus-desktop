using Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>Keeps framework-cached decorations and every rendered control on the same UI thread.</summary>
[TestClass]
public sealed class UiTestSession
{
    private static readonly Lazy<HeadlessUnitTestSession> Shared = new(() => HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder)));
    internal static HeadlessUnitTestSession Current => Shared.Value;

    [AssemblyCleanup]
    public static async Task CleanupAsync()
    {
        if (Shared.IsValueCreated) await Shared.Value.DisposeAsync();
    }
}
