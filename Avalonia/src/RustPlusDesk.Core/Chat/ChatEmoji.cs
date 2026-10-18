using System.Text.RegularExpressions;

namespace RustPlusDesk.Core.Chat;

/// <summary>A chat emoji (reference EmojiEntry): one of the custom emojis or a Rust item shown by its icon.</summary>
public sealed record ChatEmoji(string Name, string DisplayName, bool IsCustom, string Category)
{
    /// <summary>The token Rust chat understands, e.g. ":skull:" or ":rifle.ak:".</summary>
    public string Tag => $":{Name}:";
    public string ToolTip => $"{DisplayName} ({Tag})";
}

/// <summary>Part of a chat message: plain text, or an emoji token that resolved.</summary>
public sealed record ChatSegment(string Text, ChatEmoji? Emoji);

/// <summary>
/// The reference EmojiService and ChatEmojiInputHelper rules: 28 custom emojis, Rust item short names as emojis,
/// ":token:" rendering (unknown tokens stay verbatim), search (custom first, then items) and the autocomplete trigger.
/// </summary>
public sealed partial class ChatEmojiCatalog
{
    /// <summary>Reference s_rawCustomEmojis, in reference order.</summary>
    public static readonly IReadOnlyList<ChatEmoji> Custom =
    [
        new("angry", "Angry Face", true, "Face"), new("coffeecan", "Coffee Can", true, "Face"), new("cool", "Cool Shades", true, "Face"),
        new("dance", "Dancing", true, "Face"), new("exclamation", "Exclamation", true, "Emote"), new("eyebrow", "Raised Eyebrow", true, "Face"),
        new("eyes", "Looking Eyes", true, "Emote"), new("funny", "Funny Face", true, "Face"), new("happy", "Happy Face", true, "Face"),
        new("heart", "Heart", true, "Emote"), new("heartrock", "Heart Rock", true, "Item"), new("hazmat", "Hazmat", true, "Character"),
        new("trash", "Trash Bin", true, "Item"), new("bush", "Jungle Bush", true, "Item"), new("sick", "Sick Face", true, "Face"),
        new("laugh", "Laughing Face", true, "Face"), new("light", "Torch Light", true, "Item"), new("love", "Love Heart Eyes", true, "Face"),
        new("mask", "Bandana Mask", true, "Character"), new("nervous", "Nervous Sweat", true, "Face"), new("neutral", "Neutral Face", true, "Face"),
        new("scientist", "Scientist", true, "Character"), new("skull", "Skull", true, "Emote"), new("smilecry", "Smile Cry", true, "Face"),
        new("trumpet", "Trumpet", true, "Item"), new("wave", "Waving Hand", true, "Emote"), new("worried", "Worried Face", true, "Face"),
        new("yellowpin", "Yellow Pin", true, "Emote")
    ];

    public const int AutocompleteResults = 15;
    public const int PickerResults = 60;

    private static readonly Dictionary<string, ChatEmoji> CustomByName = Custom.ToDictionary(emoji => emoji.Name, StringComparer.OrdinalIgnoreCase);
    private readonly IReadOnlyList<(string ShortName, string Name)> items;
    private readonly Dictionary<string, string> itemNames;

    /// <param name="items">Rust items by short name with their display names.</param>
    public ChatEmojiCatalog(IEnumerable<(string ShortName, string Name)> items)
    {
        this.items = items.Where(item => !string.IsNullOrWhiteSpace(item.ShortName)).ToList();
        itemNames = new(StringComparer.OrdinalIgnoreCase);
        foreach (var (shortName, name) in this.items) itemNames.TryAdd(shortName, name);
    }

    [GeneratedRegex(@"\:(?<token>[a-zA-Z0-9_\.\-]+)\:")]
    private static partial Regex TokenRegex();

    /// <summary>Reference ResolveEmojiOrItem: custom emoji first, then an item short name.</summary>
    public ChatEmoji? Resolve(string token)
    {
        string clean = (token ?? "").Trim().Trim(':');
        if (clean.Length == 0) return null;
        if (CustomByName.TryGetValue(clean, out var custom)) return custom;
        return itemNames.TryGetValue(clean, out var name) ? new(clean, string.IsNullOrWhiteSpace(name) ? clean : name, false, "Item") : null;
    }

    /// <summary>Reference RichChatTextBlock: text and resolved tokens in order; unknown tokens stay as text.</summary>
    public IReadOnlyList<ChatSegment> Tokenize(string text)
    {
        var segments = new List<ChatSegment>();
        int last = 0;
        foreach (Match match in TokenRegex().Matches(text ?? ""))
        {
            if (Resolve(match.Groups["token"].Value) is not { } emoji) continue;
            if (match.Index > last) segments.Add(new(text![last..match.Index], null));
            segments.Add(new(match.Value, emoji));
            last = match.Index + match.Length;
        }
        if (last < (text?.Length ?? 0)) segments.Add(new(text![last..], null));
        return segments;
    }

    /// <summary>Reference Search: custom emojis by name or display name, then item short names; empty query lists all.</summary>
    public IReadOnlyList<ChatEmoji> Search(string? query, int max)
    {
        string clean = (query ?? "").Trim().Trim(':');
        var results = new List<ChatEmoji>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var custom in Custom)
        {
            if (clean.Length > 0 && !custom.Name.Contains(clean, StringComparison.OrdinalIgnoreCase) && !custom.DisplayName.Contains(clean, StringComparison.OrdinalIgnoreCase)) continue;
            results.Add(custom);
            seen.Add(custom.Name);
            if (results.Count >= max) return results;
        }
        foreach (var (shortName, name) in items)
        {
            if (seen.Contains(shortName) || (clean.Length > 0 && !shortName.Contains(clean, StringComparison.OrdinalIgnoreCase))) continue;
            results.Add(new(shortName, string.IsNullOrWhiteSpace(name) ? shortName : name, false, "Item"));
            seen.Add(shortName);
            if (results.Count >= max) break;
        }
        return results;
    }

    /// <summary>
    /// Reference CheckAutocomplete: the text between the last ':' before the caret and the caret, when it has no
    /// whitespace and the ':' starts a word (or follows another ':'). Returns the ':' index and the query.
    /// </summary>
    public static (int Start, string Query)? AutocompleteQuery(string text, int caret)
    {
        if (caret <= 0 || caret > text.Length) return null;
        int colon = text.LastIndexOf(':', caret - 1);
        if (colon < 0) return null;
        string query = text[(colon + 1)..caret];
        if (query.Any(char.IsWhiteSpace) || (colon > 0 && !char.IsWhiteSpace(text[colon - 1]) && text[colon - 1] != ':')) return null;
        return (colon, query);
    }

    /// <summary>Reference ApplyAutocomplete: the ":query" before the caret becomes the tag plus a space.</summary>
    public static (string Text, int Caret) ApplyAutocomplete(string text, int caret, int start, ChatEmoji emoji)
    {
        if (start < 0 || start >= text.Length) return Insert(text, caret, emoji);
        string tag = emoji.Tag + " ";
        int length = Math.Max(0, caret - start);
        string result = text.Remove(start, Math.Min(length, text.Length - start)).Insert(start, tag);
        return (result, start + tag.Length);
    }

    /// <summary>Reference InsertTag: the tag plus a space at the caret.</summary>
    public static (string Text, int Caret) Insert(string text, int caret, ChatEmoji emoji)
    {
        int at = Math.Clamp(caret, 0, text.Length);
        string tag = emoji.Tag + " ";
        return (text.Insert(at, tag), at + tag.Length);
    }
}
