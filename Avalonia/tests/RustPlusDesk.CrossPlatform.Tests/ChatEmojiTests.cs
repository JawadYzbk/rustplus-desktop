using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ChatEmojiTests
{
    private static ChatEmojiCatalog Catalog() => new([("rifle.ak", "Assault Rifle"), ("rock", "Rock"), ("skull.human", "Human Skull"), ("", "Nameless")]);

    [TestMethod]
    public void Catalog_ResolvesTokenizesAndSearchesLikeTheReference()
    {
        var catalog = Catalog();
        Assert.AreEqual(28, ChatEmojiCatalog.Custom.Count);
        Assert.AreEqual(":skull:", catalog.Resolve(":SKULL:")!.Tag);
        Assert.IsTrue(catalog.Resolve("skull")!.IsCustom);
        var item = catalog.Resolve("rifle.ak")!;
        Assert.IsFalse(item.IsCustom);
        Assert.AreEqual("Assault Rifle (:rifle.ak:)", item.ToolTip);
        Assert.IsNull(catalog.Resolve("nope"));

        var segments = catalog.Tokenize("gg :skull: at :nope: with :rifle.ak:!");
        CollectionAssert.AreEqual(new[] { "gg ", ":skull:", " at :nope: with ", ":rifle.ak:", "!" }, segments.Select(segment => segment.Text).ToArray(),
            "Unknown tokens stay as typed.");
        Assert.AreEqual("rifle.ak", segments[3].Emoji!.Name);
        Assert.AreEqual(0, catalog.Tokenize("").Count);
        Assert.AreEqual("plain", catalog.Tokenize("plain").Single().Text);

        var skull = catalog.Search(":sku", 30);
        CollectionAssert.AreEqual(new[] { "skull", "skull.human" }, skull.Select(emoji => emoji.Name).ToArray(), "Custom emojis first, then items.");
        Assert.AreEqual("smilecry", catalog.Search("Smile Cry", 30).Single().Name, "Display names match for custom emojis.");
        Assert.AreEqual(31, catalog.Search("", 100).Count, "Empty query lists everything.");
        Assert.AreEqual(15, catalog.Search("", 15).Count);
    }

    [TestMethod]
    public void Autocomplete_TriggersOnlyAtAWordStart_AndReplacesTheQuery()
    {
        Assert.AreEqual((3, "sk"), ChatEmojiCatalog.AutocompleteQuery("hi :sk", 6));
        Assert.AreEqual((0, ""), ChatEmojiCatalog.AutocompleteQuery(":", 1));
        Assert.AreEqual((7, "a"), ChatEmojiCatalog.AutocompleteQuery(":skull::a", 9), "A ':' right after another starts a new token.");
        Assert.IsNull(ChatEmojiCatalog.AutocompleteQuery("time 10:30", 10), "Inside a word it is not a token.");
        Assert.IsNull(ChatEmojiCatalog.AutocompleteQuery("hi :sk ull", 10), "Whitespace ends the query.");
        Assert.IsNull(ChatEmojiCatalog.AutocompleteQuery("hello", 0));
        var skull = ChatEmojiCatalog.Custom.Single(emoji => emoji.Name == "skull");
        Assert.AreEqual(("hi :skull: ", 11), ChatEmojiCatalog.ApplyAutocomplete("hi :sk", 6, 3, skull));
        Assert.AreEqual(("hi :skull: there", 11), ChatEmojiCatalog.ApplyAutocomplete("hi :skthere", 6, 3, skull));
        Assert.AreEqual(("a:skull: b", 9), ChatEmojiCatalog.Insert("ab", 1, skull));
        Assert.AreEqual(("ab:skull: ", 10), ChatEmojiCatalog.ApplyAutocomplete("ab", 2, -1, skull), "Without a query the tag is inserted.");
    }

    [TestMethod]
    public void ViewModel_DrivesSuggestionsPickerAndTheLimit()
    {
        var model = new ChatViewModel { Emoji = Catalog() };
        model.Draft = "raid :hum";
        model.UpdateEmojiAutocomplete(9);
        Assert.IsTrue(model.IsEmojiSuggestionsOpen);
        Assert.AreEqual("skull.human", model.SelectedEmojiSuggestion!.Name);
        model.Draft = "raid :sku";
        model.UpdateEmojiAutocomplete(9);
        Assert.AreEqual("skull", model.SelectedEmojiSuggestion!.Name);
        model.MoveEmojiSuggestion(1);
        Assert.AreEqual("skull.human", model.SelectedEmojiSuggestion.Name);
        model.MoveEmojiSuggestion(1);
        Assert.AreEqual("skull", model.SelectedEmojiSuggestion.Name, "The selection wraps.");
        Assert.AreEqual(13, model.AcceptEmojiSuggestion(9));
        Assert.AreEqual("raid :skull: ", model.Draft);
        Assert.IsFalse(model.IsEmojiSuggestionsOpen);
        model.Draft = "x :zzz";
        model.UpdateEmojiAutocomplete(6);
        Assert.IsFalse(model.IsEmojiSuggestionsOpen, "No matches, no list.");

        CollectionAssert.AreEqual(ChatEmojiCatalog.Custom.ToArray(), model.EmojiPickerItems.ToArray(), "The picker opens on the custom emojis.");
        model.EmojiPickerQuery = "rock";
        CollectionAssert.AreEqual(new[] { "heartrock", "rock" }, model.EmojiPickerItems.Select(emoji => emoji.Name).ToArray());
        model.EmojiPickerQuery = "nothing";
        Assert.IsTrue(model.HasNoEmojiPickerItems);
        model.Draft = "gg";
        Assert.AreEqual(8, model.InsertEmoji(ChatEmojiCatalog.Custom[0], 0));
        Assert.AreEqual(":angry: gg", model.Draft);
        model.Draft = new string('x', 125);
        Assert.IsNull(model.InsertEmoji(ChatEmojiCatalog.Custom[0], 125));
        Assert.AreEqual(125, model.Draft.Length);
        Assert.AreEqual("That emoji does not fit in the 128-character limit.", model.Notice);
    }

    [TestMethod]
    public async Task ChatView_DrawsEmojiInline_AndCompletesTokensFromTheKeyboard()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var session = new ChatTests.ChatSession { HasClan = false };
            session.History[ChatChannel.Team] = [new(ChatChannel.Team, 7, "Ann", "raid at G12 :skull: bring :rifle.ak: :unknown:", DateTimeOffset.UtcNow)];
            var model = new ChatViewModel { Emoji = new(RustPlusDesk.Infrastructure.Items.BundledItemCatalog.Icons.Values.Select(item => (item.ShortName, item.Name))) };
            using var cts = new CancellationTokenSource();
            var running = model.RunAsync(session, cts.Token);
            var window = new Window { Width = 700, Height = 300, Content = new ChatView { DataContext = model } };
            try
            {
                window.Show();
                await CompanionListenerViewModelTests.Eventually(() => model.Lines.Count == 1);
                window.UpdateLayout();
                var line = window.GetVisualDescendants().OfType<ChatMessageText>().Single();
                var pictures = line.Inlines!.OfType<InlineUIContainer>().ToList();
                Assert.AreEqual(2, pictures.Count, "The skull emoji and the rifle icon are drawn; :unknown: stays text.");
                Assert.IsInstanceOfType<Image>(pictures[0].Child);
                Assert.IsNotNull(((Image)pictures[0].Child!).Source, "The bundled thumbnail loads.");
                StringAssert.Contains(string.Concat(line.Inlines!.OfType<Run>().Select(run => run.Text)), ":unknown:");
                Assert.IsNotNull(line.ContextMenu, "The message menu is kept.");

                var draft = window.GetVisualDescendants().OfType<TextBox>().Single(box => AutomationProperties.GetName(box) == "Chat message");
                draft.Focus();
                draft.Text = "gg :sku";
                draft.CaretIndex = 7;
                model.UpdateEmojiAutocomplete(7);
                Assert.IsTrue(model.IsEmojiSuggestionsOpen);
                var view = window.GetVisualDescendants().OfType<ChatView>().Single();
                view.DraftKeyDown(draft, new KeyEventArgs { Key = Key.Tab, RoutedEvent = InputElement.KeyDownEvent });
                Assert.AreEqual("gg :skull: ", model.Draft);
                Assert.AreEqual(11, draft.CaretIndex);
                window.UpdateLayout();
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-chat-emoji.png"), PngBitmapEncoderOptions.Default);
                }
                var picker = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Insert emoji");
                Assert.IsNotNull(picker.Flyout);
            }
            finally { cts.Cancel(); await running; window.Close(); }
            return true;
        }, default);
    }
}
