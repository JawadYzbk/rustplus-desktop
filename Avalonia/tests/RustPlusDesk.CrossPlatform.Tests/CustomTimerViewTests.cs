using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.LogicalTree;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference BtnCustomTimer_Click / PopupCustomTimer: the map toolbar's timer button opens the popup with the active
/// timers and the add form, and the popup reads the selected profile's timers.
/// </summary>
[TestClass]
public sealed class CustomTimerViewTests
{
    [TestMethod]
    public async Task TimerFlyout_OpensFromTheToolbar_AndRendersTheAddForm()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), fixture.Path, new UnusedSessionFactory());
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark };
            try
            {
                await model.RefreshCommand.ExecuteAsync(null);
                window.Show();
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);

                var button = window.GetVisualDescendants().OfType<Button>().Single(candidate => candidate.Name == "TimersButton");
                var flyout = (Flyout)button.Flyout!;
                flyout.ShowAt(button);
                using (var opened = window.CaptureRenderedFrame()) Assert.IsNotNull(opened);

                Assert.IsTrue(flyout.IsOpen);
                var content = (StackPanel)flyout.Content!;
                Assert.AreSame(model, content.DataContext, "The popup reads the workspace's selected profile.");
                Assert.AreEqual("Active Timers", ((TextBlock)content.Children[0]).Text);
                Assert.AreEqual("No active timers.", ((TextBlock)content.Children[1]).Text);
                var list = (ItemsControl)content.Children[2];
                Assert.AreSame(model.Timers.Timers, list.ItemsSource);

                var checkBox = content.GetLogicalDescendants().OfType<CheckBox>()
                    .Single(box => Equals(box.Content, "Send timer alerts to chat"));
                Assert.IsTrue(checkBox.IsChecked);
                checkBox.IsChecked = false;
                Assert.IsFalse(model.Timers.AlertCustomTimer, "The popup's alert switch is the profile's AlertCustomTimer.");

                Assert.IsNotNull(content.GetLogicalDescendants().OfType<Button>()
                    .Single(candidate => Equals(candidate.Content, "Add Timer")));
                using (var toggled = window.CaptureRenderedFrame()) Assert.IsNotNull(toggled);
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, default);
    }

    private sealed class Picker : IProfileFilePicker { public Task<string?> PickAsync() => Task.FromResult<string?>(null); }
}
