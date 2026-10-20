using System.Globalization;

namespace RustPlusDesk.Core.Localization;

/// <summary>One language the app ships, as the settings picker sees it.</summary>
public sealed record AppLanguage(string Code, string Name);

/// <summary>
/// The languages the app ships, in one place (reference Helpers.AppLanguages).
///
/// The codes match the shipped resource catalogues. The names are in the language itself: somebody looking for their own
/// language reads it in their own language, not in English.
/// </summary>
public static class AppLanguages
{
    public static IReadOnlyList<AppLanguage> All { get; } =
    [
        new("af-ZA", "Afrikaans"),
        new("ar-SA", "العربية"),
        new("ca-ES", "Català"),
        new("cs-CZ", "Čeština"),
        new("da-DK", "Dansk"),
        new("de-DE", "Deutsch"),
        new("el-GR", "Ελληνικά"),
        new("en-US", "English"),
        new("es-ES", "Español"),
        new("fi-FI", "Suomi"),
        new("fr-FR", "Français"),
        new("he-IL", "עברית"),
        new("hu-HU", "Magyar"),
        new("it-IT", "Italiano"),
        new("ja-JP", "日本語"),
        new("ko-KR", "한국어"),
        new("nl-NL", "Nederlands"),
        new("no-NO", "Norsk"),
        new("pl-PL", "Polski"),
        new("pt-BR", "Português (BR)"),
        new("pt-PT", "Português (PT)"),
        new("ro-RO", "Română"),
        new("ru-RU", "Русский"),
        new("sr-Latn-RS", "Srpski"),
        new("sv-SE", "Svenska"),
        new("tr-TR", "Türkçe"),
        new("uk-UA", "Українська"),
        new("vi-VN", "Tiếng Việt"),
        new("zh-CN", "简体中文"),
        new("zh-Hans", "简体中文 (Hans)"),
        new("zh-Hant", "繁體中文 (Hant)"),
        new("zh-TW", "繁體中文")
    ];

    /// <summary>
    /// The shipped code closest to a culture, or null when the app ships nothing for it. A Windows set to de-AT
    /// resolves to de-DE rather than to nothing: the point of reporting a language is being findable by people who
    /// speak it, and Austrian German finds German.
    /// </summary>
    public static string? Resolve(CultureInfo? culture)
    {
        string? name = culture?.Name;
        if (string.IsNullOrWhiteSpace(name)) return null;

        var exact = All.FirstOrDefault(language => string.Equals(language.Code, name, StringComparison.OrdinalIgnoreCase));
        if (exact is not null) return exact.Code;

        string part = name.Split('-')[0];

        // Bokmål and Nynorsk are what Windows reports; the catalogue is the older two-letter code.
        if (part.Equals("nb", StringComparison.OrdinalIgnoreCase)
            || part.Equals("nn", StringComparison.OrdinalIgnoreCase))
            return "no-NO";

        return All.FirstOrDefault(language =>
            language.Code.Split('-')[0].Equals(part, StringComparison.OrdinalIgnoreCase))?.Code;
    }

    /// <summary>The culture for a shipped language code, or null when the code is not one of ours.</summary>
    public static CultureInfo? CultureFor(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return null;
        if (!All.Any(language => string.Equals(language.Code, code, StringComparison.OrdinalIgnoreCase))) return null;
        try { return CultureInfo.GetCultureInfo(code); }
        catch (CultureNotFoundException) { return null; }
    }
}
