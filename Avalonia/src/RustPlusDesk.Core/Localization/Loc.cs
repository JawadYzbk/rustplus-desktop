using System.Globalization;
using System.Resources;

namespace RustPlusDesk.Core.Localization;

/// <summary>
/// Reference Helpers.Loc: a string is looked up in the language the user chose, not in the operating system's
/// language, and the key itself is never shown to the user.
///
/// The resource manager is set by the app at start-up, because the strings live in the app's assembly and the core
/// does not know about it.
/// </summary>
public static class Loc
{
    private static ResourceManager? manager;

    /// <summary>The app's resource catalogue (the reference's Resources.resx and its satellites).</summary>
    public static ResourceManager? Manager
    {
        get => manager;
        set => manager = value;
    }

    /// <summary>The language the strings are read in (the reference's Resources.Culture).</summary>
    public static CultureInfo? Culture { get; set; }

    /// <summary>Null rather than the key when a string is missing, for callers with their own default.</summary>
    public static string? TextOrNull(string key)
    {
        if (manager is null) return null;
        string? value;
        try { value = manager.GetString(key, Culture); }
        catch (MissingManifestResourceException) { return null; }
        return string.IsNullOrWhiteSpace(value) || value == key ? null : value;
    }

    /// <summary>The string for the key, or the fallback when there is none.</summary>
    public static string Text(string key, string fallback) => TextOrNull(key) ?? fallback;

    /// <summary>
    /// Sets the language the strings are read in. The thread's UI culture follows, so anything that asks the
    /// operating system for a language gets the one the user chose.
    /// </summary>
    public static void UseCulture(CultureInfo? culture)
    {
        Culture = culture;
        if (culture is null) return;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
        CultureInfo.CurrentUICulture = culture;
    }
}
