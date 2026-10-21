using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MainWindow.Map.Chat.QuickCommandChips: the commands the profile answers, offered as chips that fill the
/// chat input.
/// </summary>
[TestClass]
public sealed class QuickCommandChipTests
{
    [TestMethod]
    public void TheChipsAreBuiltFromTheProfileSettings()
    {
        var model = new ChatViewModel
        {
            CommandSettings = ChatCommandSettings.Defaults with { Prefix = "!" },
        };

        Assert.IsTrue(model.QuickCommandChips.Count > 0, "The chips are built.");
        Assert.IsTrue(model.QuickCommandChips.All(chip => chip.Command.StartsWith("!", StringComparison.Ordinal)),
            "Every chip carries the profile's prefix.");
        Assert.IsTrue(model.QuickCommandChips.Any(chip => chip.Command == "!cargo"), "The cargo command is offered.");
        Assert.IsTrue(model.QuickCommandChips.Any(chip => chip.Command == "!code"), "The base codes are offered to the team.");
    }

    [TestMethod]
    public void ThePrefixIsTheProfilesOwn()
    {
        var model = new ChatViewModel
        {
            CommandSettings = ChatCommandSettings.Defaults with { Prefix = "." },
        };

        Assert.IsTrue(model.QuickCommandChips.All(chip => chip.Command.StartsWith(".", StringComparison.Ordinal)),
            "The chips carry the prefix the profile uses.");
    }

    [TestMethod]
    public void TheClanIsOfferedTheTimeAndTheTeamTheBaseCodes()
    {
        var model = new ChatViewModel
        {
            CommandSettings = ChatCommandSettings.Defaults,
        };

        Assert.IsTrue(model.QuickCommandChips.Any(chip => chip.Command == "!code"), "The team is offered the base codes.");
        Assert.IsFalse(model.QuickCommandChips.Any(chip => chip.Command == "!time"), "The team is not offered the time.");

        model.IsClan = true;

        Assert.IsTrue(model.QuickCommandChips.Any(chip => chip.Command == "!time"), "The clan is offered the time.");
        Assert.IsFalse(model.QuickCommandChips.Any(chip => chip.Command == "!code"), "The clan is not offered the base codes.");
    }

    [TestMethod]
    public void TheTimerChipTypesThePairTheTimerWants()
    {
        var model = new ChatViewModel
        {
            CommandSettings = ChatCommandSettings.Defaults,
        };

        var timer = model.QuickCommandChips.FirstOrDefault(chip => chip.Label.Contains("15", StringComparison.Ordinal));
        Assert.IsNotNull(timer, "The timer chip is there.");
        Assert.IsTrue(timer!.Command.EndsWith("oilrig,15", StringComparison.Ordinal),
            "The timer is a name and a number in one pair, as the timer wants it.");
    }

    [TestMethod]
    public void AChosenChipFillsTheInput()
    {
        var model = new ChatViewModel
        {
            CommandSettings = ChatCommandSettings.Defaults,
        };

        var chip = model.QuickCommandChips.First();
        model.UseQuickCommand(chip);

        Assert.AreEqual(chip.Command, model.Draft, "The chip fills the input, and the player sends it.");
    }

    [TestMethod]
    public void ThereAreNoChipsWithoutAProfile()
    {
        var model = new ChatViewModel();

        Assert.AreEqual(0, model.QuickCommandChips.Count, "Without a profile there is nothing to offer.");
    }
}
