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

    /// <summary>
    /// Raised when the language changes, so a view can read its own labels again.
    ///
    /// The subscribers are held weakly: a view model that is thrown away does not stay in the list, so a language
    /// change never reaches a dead view and a long session does not pile up handlers.
    /// </summary>
    public static event EventHandler? LanguageChanged
    {
        add
        {
            if (value is null) return;
            lock (Handlers)
                Handlers.Add(new Subscription(value));
        }
        remove
        {
            if (value is null) return;
            lock (Handlers)
                Handlers.RemoveAll(entry => entry.Matches(value));
        }
    }

    private static readonly List<Subscription> Handlers = new();

    /// <summary>One subscriber, held without keeping the object it belongs to alive.</summary>
    private sealed class Subscription
    {
        private readonly WeakReference? target;
        private readonly System.Reflection.MethodInfo? method;
        private readonly EventHandler? handler;

        public Subscription(EventHandler handler)
        {
            // A handler that belongs to no object (a static method) has nothing to keep alive, so it is held as it is.
            if (handler.Target is null) { this.handler = handler; return; }
            target = new WeakReference(handler.Target);
            method = handler.Method;
        }

        public bool Matches(EventHandler other) =>
            handler is not null ? handler == other : other.Target is not null && method == other.Method && target is not null
                && target.IsAlive && ReferenceEquals(target.Target, other.Target);

        /// <summary>The handler to call, or null when the object it belongs to is gone.</summary>
        public EventHandler? Live()
        {
            if (handler is not null) return handler;
            if (target is null || method is null || !target.IsAlive || target.Target is not { } alive) return null;
            try { return (EventHandler)method.CreateDelegate(typeof(EventHandler), alive); }
            catch (System.ArgumentException) { return null; }
            catch (System.MissingMethodException) { return null; }
        }
    }

    /// <summary>Tells the subscribers that the language changed.</summary>
    private static void RaiseLanguageChanged()
    {
        List<EventHandler> live = new();
        lock (Handlers)
        {
            for (int index = Handlers.Count - 1; index >= 0; index--)
            {
                EventHandler? handler = Handlers[index].Live();
                if (handler is not null) { live.Add(handler); continue; }
                Handlers.RemoveAt(index);
            }
        }
        foreach (var handler in live) handler(null, EventArgs.Empty);
    }

    /// <summary>Null rather than the key when a string is missing, for callers with their own default.</summary>
    public static string? TextOrNull(string key)
    {
        if (manager is null) return null;
        string? value;
        try { value = manager.GetString(key, Culture); }
        catch (MissingManifestResourceException) { return null; }
        // The reference's generated resources hand back the key itself when a string is missing, so it compares the
        // two. This port reads the catalogue through the resource manager, which hands back null instead - and a
        // string whose own wording is its key ("Clear", "Reset", "Save") is a real string, not a missing one.
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    /// <summary>The string for the key, or the fallback when there is none.</summary>
    public static string Text(string key, string fallback) => TextOrNull(key) ?? fallback;

    /// <summary>
    /// Sets the language the strings are read in. The thread's UI culture follows, so anything that asks the
    /// operating system for a language gets the one the user chose.
    ///
    /// A caller that only wants the strings read in another language - a test that reads them in German, say - can
    /// leave the process's culture alone: the language of the app is read from <see cref="Culture"/>, not from the
    /// thread.
    /// </summary>
    public static void UseCulture(CultureInfo? culture, bool applyToThread = true)
    {
        Culture = culture;
        if (culture is not null && applyToThread)
        {
            CultureInfo.DefaultThreadCurrentUICulture = culture;
            CultureInfo.CurrentUICulture = culture;
        }
        RaiseLanguageChanged();
    }
}
