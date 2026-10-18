using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using RustPlusDesk.Avalonia.Shared.Controls;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.Avalonia.Features.Chat;

/// <summary>
/// A chat line (reference RichChatTextBlock): time, author and the message, with ":emoji:" and ":item.shortname:"
/// tokens drawn inline as pictures; unknown tokens stay as typed.
/// </summary>
public sealed class ChatMessageText : TextBlock
{
    public const double EmojiSize = 18;
    public static readonly StyledProperty<ChatLine?> LineProperty = AvaloniaProperty.Register<ChatMessageText, ChatLine?>(nameof(Line));
    public static readonly StyledProperty<ChatEmojiCatalog?> CatalogProperty = AvaloniaProperty.Register<ChatMessageText, ChatEmojiCatalog?>(nameof(Catalog));
    private static readonly Dictionary<string, Bitmap?> CustomImages = new(StringComparer.OrdinalIgnoreCase);

    public ChatLine? Line { get => GetValue(LineProperty); set => SetValue(LineProperty, value); }
    public ChatEmojiCatalog? Catalog { get => GetValue(CatalogProperty); set => SetValue(CatalogProperty, value); }

    protected override Type StyleKeyOverride => typeof(TextBlock);

    static ChatMessageText()
    {
        LineProperty.Changed.AddClassHandler<ChatMessageText>((text, _) => text.Rebuild());
        CatalogProperty.Changed.AddClassHandler<ChatMessageText>((text, _) => text.Rebuild());
    }

    private void Rebuild()
    {
        var inlines = new InlineCollection();
        if (Line is { } line)
        {
            var time = new Run(line.Time) { FontSize = 11 };
            time.Bind(TextElement.ForegroundProperty, this.GetResourceObservable("TextSubtle"));
            var name = new Run(line.Name) { FontWeight = FontWeight.SemiBold };
            name.Bind(TextElement.ForegroundProperty, this.GetResourceObservable("Accent"));
            inlines.Add(time);
            inlines.Add(new Run(" "));
            inlines.Add(name);
            inlines.Add(new Run(": "));
            foreach (var segment in Catalog?.Tokenize(line.Text) ?? [new ChatSegment(line.Text, null)])
                inlines.Add(segment.Emoji is { } emoji && Visual(emoji) is { } visual ? new InlineUIContainer(visual) { BaselineAlignment = BaselineAlignment.Center } : new Run(segment.Text));
        }
        Inlines = inlines;
    }

    /// <summary>The picture for an emoji: the bundled thumbnail for custom ones, the item icon for items.</summary>
    public static Control? Visual(ChatEmoji emoji)
    {
        Control? visual = emoji.IsCustom
            ? CustomImage(emoji.Name) is { } bitmap ? new Image { Source = bitmap } : null
            : new ItemIcon { ShortName = emoji.Name };
        if (visual is null) return null;
        visual.Width = visual.Height = EmojiSize;
        visual.Margin = new Thickness(2, 0);
        visual.VerticalAlignment = VerticalAlignment.Center;
        ToolTip.SetTip(visual, emoji.ToolTip);
        return visual;
    }

    public static Bitmap? CustomImage(string name)
    {
        if (CustomImages.TryGetValue(name, out var cached)) return cached;
        Bitmap? bitmap = null;
        try
        {
            var uri = new Uri($"avares://RustPlusDesk.Avalonia/Assets/Emojis/{name}.png");
            if (AssetLoader.Exists(uri)) using (var stream = AssetLoader.Open(uri)) bitmap = new Bitmap(stream);
        }
        catch (Exception) { }
        return CustomImages[name] = bitmap;
    }
}

/// <summary>Binding helper: custom emoji name → bundled thumbnail.</summary>
public static class ChatEmojiImages
{
    public static readonly global::Avalonia.Data.Converters.IValueConverter Custom =
        new global::Avalonia.Data.Converters.FuncValueConverter<string?, Bitmap?>(name => name is null ? null : ChatMessageText.CustomImage(name));
}
