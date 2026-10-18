using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Shared.Icons;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DesignSystemTests
{
    /// <summary>Icon styles in Components.axaml (icon/subtle buttons, chips) must reach the Icon control.</summary>
    [TestMethod]
    public async Task IconStyles_ApplyInsideIconButtonsAndChips() => await UiTestSession.Current.Dispatch(async () =>
    {
        var plain = new Icon { Symbol = FluentSymbol.Alert24 };
        var chipIcon = new Icon { Symbol = FluentSymbol.Search24 };
        var button = new Button { Content = plain };
        button.Classes.Add("icon");
        var chip = new ToggleButton { Content = chipIcon, IsChecked = true };
        chip.Classes.Add("chip");
        var window = new Window { Content = new StackPanel { Children = { button, chip } } };
        try
        {
            window.Show();
            window.UpdateLayout();
            window.TryFindResource("TextSubtle", window.ActualThemeVariant, out var subtle);
            window.TryFindResource("Accent", window.ActualThemeVariant, out var accent);
            Assert.AreSame(subtle, plain.Foreground, "Button.icon icons|Icon uses the subtle text colour.");
            Assert.AreSame(accent, chipIcon.Foreground, "A checked chip's icon uses the accent.");
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);
}
