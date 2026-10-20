using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Shell;
using RustPlusDesk.Core.Platform;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference SetupTrayIcon in App.xaml.cs: the system tray of the reference app, with the same menu, the same
/// tooltip and the same behaviour when the user clicks the icon.
///
/// The tray icon itself is Avalonia's, which needs a desktop; the tests therefore work on the state behind the
/// interface and never ask for a real icon.
/// </summary>
[TestClass]
public sealed class TrayIconTests
{
    /// <summary>The app's own tray icon, without the platform: it remembers what it was told and can play it back.</summary>
    private sealed class FakeTrayIcon : ITrayIcon
    {
        public bool IsVisible { get; set; }
        public string? Tooltip { get; set; }
        public byte[]? Icon { get; set; }
        public IReadOnlyList<TrayMenuItem> Menu { get; private set; } = [];
        public (string Title, string Message)? Balloon { get; private set; }
        public bool Disposed { get; private set; }

        public void ShowBalloon(string title, string message) => Balloon = (title, message);

        public void ShowContextMenu(IReadOnlyList<TrayMenuItem> items) => Menu = items;

        public event EventHandler? Clicked;
        public event EventHandler? ContextMenuRequested;
        public event EventHandler<TrayMenuItem>? MenuItemInvoked;

        /// <summary>The user clicked the icon.</summary>
        public void RaiseClicked() => Clicked?.Invoke(this, EventArgs.Empty);

        /// <summary>The platform is about to show the menu.</summary>
        public void RaiseContextMenuRequested() => ContextMenuRequested?.Invoke(this, EventArgs.Empty);

        /// <summary>The user chose a line of the menu.</summary>
        public void Choose(TrayMenuItem item) => MenuItemInvoked?.Invoke(this, item);

        public void Dispose() => Disposed = true;
    }

    // ---- the menu ---------------------------------------------------------------------------

    [TestMethod]
    public void TheMenuIsTheReferencesMenu()
    {
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => false);

        Assert.AreEqual(6, tray.Menu.Count, "The reference's menu has six lines.");
        Assert.AreEqual("Tracking: Idle", tray.Menu[0].Header, "The reference's status line.");
        Assert.IsFalse(tray.Menu[0].IsEnabled, "The status line is there to be read, not chosen.");
        Assert.AreEqual("Last update: --:--:--", tray.Menu[1].Header, "The reference's last pull, with no pull yet.");
        Assert.IsFalse(tray.Menu[1].IsEnabled, "The last pull is there to be read, not chosen.");
        Assert.IsTrue(tray.Menu[2].IsSeparator, "A separator between the status lines and the actions.");
        Assert.AreEqual("Open Rust+ Desk", tray.Menu[3].Header, "The reference's own wording.");
        Assert.AreEqual(TrayIconController.OpenId, tray.Menu[3].Id, "The line opens the app.");
        Assert.IsTrue(tray.Menu[4].IsSeparator, "A separator before the line that leaves.");
        Assert.AreEqual("Exit", tray.Menu[5].Header);
        Assert.AreEqual(TrayIconController.ExitId, tray.Menu[5].Id, "The line leaves the app.");
    }

    [TestMethod]
    public void TheMenuSaysWhatTheAppIsDoing()
    {
        bool tracking = false;
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => tracking);

        Assert.AreEqual("Tracking: Idle", tray.Menu[0].Header, "The app is not following a server yet.");

        tracking = true;
        controller.Refresh();
        Assert.AreEqual("Tracking: Active", tray.Menu[0].Header, "The reference says Active while it tracks.");
    }

    [TestMethod]
    public void TheMenuSaysWhenTheAppLastHeardFromTheServer()
    {
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => true);

        controller.NoteServerUpdate(new DateTime(2026, 10, 20, 13, 45, 5));

        Assert.AreEqual("Last update: 13:45:05", tray.Menu[1].Header, "The reference shows the time of its last pull.");
    }

    [TestMethod]
    public void TheMenuIsFilledAgainWhenThePlatformAsksForIt()
    {
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => false);

        tray.RaiseContextMenuRequested();

        Assert.AreEqual(6, tray.Menu.Count, "The menu is filled with what is true now.");
    }

    // ---- what the user chooses ---------------------------------------------------------------

    [TestMethod]
    public void ChoosingOpenTellsTheShell()
    {
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => false);
        bool opened = false;
        controller.OpenRequested += (_, _) => opened = true;

        tray.Choose(tray.Menu[3]);

        Assert.IsTrue(opened, "The shell is asked to open the app.");
    }

    [TestMethod]
    public void ChoosingExitTellsTheShell()
    {
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => false);
        bool left = false;
        controller.ExitRequested += (_, _) => left = true;

        tray.Choose(tray.Menu[5]);

        Assert.IsTrue(left, "The shell is asked to close the app.");
    }

    [TestMethod]
    public void ADoubleClickOpensTheApp()
    {
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => false);
        bool opened = false;
        controller.OpenRequested += (_, _) => opened = true;

        tray.RaiseClicked();
        Assert.IsFalse(opened, "The reference does nothing on a single click.");

        tray.RaiseClicked();
        Assert.IsTrue(opened, "The reference opens the app on a double click.");
    }

    // ---- the icon ----------------------------------------------------------------------------

    [TestMethod]
    public void TheIconIsSet()
    {
        var tray = new FakeTrayIcon();
        byte[] icon = [1, 2, 3];
        using var controller = new TrayIconController(tray, () => false, icon);

        CollectionAssert.AreEqual(icon, tray.Icon, "The tray shows the app's own icon.");
        Assert.IsTrue(tray.IsVisible, "The reference shows its tray icon as soon as the app starts.");
    }

    [TestMethod]
    public void TheTooltipSaysWhatTheAppIsDoing()
    {
        bool tracking = false;
        var tray = new FakeTrayIcon();
        using var controller = new TrayIconController(tray, () => tracking);

        Assert.AreEqual("Rust+ Desk Tracker", tray.Tooltip, "The reference's own tooltip while it does not track.");

        tracking = true;
        controller.Refresh();
        Assert.AreEqual("Rust+ Desk (Tracking --:--:--)", tray.Tooltip, "The reference's tooltip while it tracks.");

        controller.NoteServerUpdate(new DateTime(2026, 10, 20, 13, 45, 5));
        Assert.AreEqual("Rust+ Desk (Tracking 13:45:05)", tray.Tooltip, "The tooltip carries the time of the last pull.");
    }

    // ---- the controller ----------------------------------------------------------------------

    [TestMethod]
    public void TheControllerCanBeMadeAndThrownAwayWithoutADesktop()
    {
        var tray = new FakeTrayIcon();
        var controller = new TrayIconController(tray, () => false);

        controller.Dispose();

        Assert.IsTrue(tray.Disposed, "The controller lets the tray icon go.");
    }

    [TestMethod]
    public void AThrownAwayControllerStopsListening()
    {
        var tray = new FakeTrayIcon();
        var controller = new TrayIconController(tray, () => false);
        bool opened = false;
        controller.OpenRequested += (_, _) => opened = true;

        controller.Dispose();
        tray.RaiseClicked();
        tray.RaiseClicked();

        Assert.IsFalse(opened, "A controller that is gone is not asked any more.");
    }
}
