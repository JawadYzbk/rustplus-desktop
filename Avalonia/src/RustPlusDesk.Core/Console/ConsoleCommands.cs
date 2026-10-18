using System.Text.Json.Serialization;

namespace RustPlusDesk.Core.Console;

/// <summary>One editable value of a command template (reference ConsoleCommandParam), e.g. {item} in "craft.add {item} {amount}".</summary>
public sealed record ConsoleCommandParamSpec
{
    [JsonPropertyName("name")] public string Name { get; init; } = "";
    /// <summary>text, int, float, bool or item.</summary>
    [JsonPropertyName("type")] public string Type { get; init; } = "text";
    [JsonPropertyName("label")] public string Label { get; init; } = "";
    [JsonPropertyName("default")] public string Default { get; init; } = "";
    [JsonPropertyName("min")] public double? Min { get; init; }
    [JsonPropertyName("max")] public double? Max { get; init; }
}

/// <summary>Title and description in one language.</summary>
public sealed record ConsoleCommandText
{
    [JsonPropertyName("title")] public string? Title { get; init; }
    [JsonPropertyName("description")] public string? Description { get; init; }
}

/// <summary>A command as shipped in the reference console-commands.json.</summary>
public sealed record ConsoleCommandSpec
{
    [JsonPropertyName("id")] public string Id { get; init; } = "";
    /// <summary>client or admin (the two reference tabs).</summary>
    [JsonPropertyName("category")] public string Category { get; init; } = "client";
    [JsonPropertyName("group")] public string Group { get; init; } = "";
    [JsonPropertyName("featured")] public bool Featured { get; init; }
    [JsonPropertyName("title")] public string Title { get; init; } = "";
    [JsonPropertyName("description")] public string Description { get; init; } = "";
    [JsonPropertyName("i18n")] public Dictionary<string, ConsoleCommandText>? Translations { get; init; }
    [JsonPropertyName("command")] public string Command { get; init; } = "";
    [JsonPropertyName("bindable")] public bool Bindable { get; init; }
    /// <summary>The command carries its own quoting (meta.exec) and is bound verbatim.</summary>
    [JsonPropertyName("rawBind")] public bool RawBind { get; init; }
    [JsonPropertyName("defaultBind")] public string? DefaultBind { get; init; }
    [JsonPropertyName("params")] public List<ConsoleCommandParamSpec> Params { get; init; } = [];
}

public sealed record ConsoleCommandCatalog
{
    [JsonPropertyName("version")] public int Version { get; init; }
    [JsonPropertyName("commands")] public List<ConsoleCommandSpec> Commands { get; init; } = [];
}

/// <summary>What the reference persists per command in console-helper.json: only what differs from the catalogue.</summary>
public sealed record ConsoleCommandUserState
{
    [JsonPropertyName("bind")] public string? Bind { get; init; }
    [JsonPropertyName("values")] public Dictionary<string, string>? Values { get; init; }
}

/// <summary>Shipped catalogue plus the user's binds and values, stored apart so catalogue updates never wipe binds.</summary>
public interface IConsoleCommandStore
{
    ConsoleCommandCatalog LoadCatalog();
    Task<IReadOnlyDictionary<string, ConsoleCommandUserState>> LoadStateAsync(CancellationToken ct);
    Task SaveStateAsync(IReadOnlyDictionary<string, ConsoleCommandUserState> state, CancellationToken ct);
}

/// <summary>Reference ConsoleCommandDef / ConsoleCommandLibrary rules.</summary>
public static class ConsoleCommandRules
{
    public const string Client = "client";
    public const string Admin = "admin";

    /// <summary>Reference ResolvedCommand: every {name} placeholder replaced by its value.</summary>
    public static string Resolve(string template, IEnumerable<KeyValuePair<string, string>> values)
    {
        string text = template;
        foreach (var (name, value) in values) text = text.Replace("{" + name + "}", value);
        return text;
    }

    /// <summary>
    /// Reference BindLine: "bind key command"; a command with a separator or space is wrapped in quotes (inner quotes
    /// become apostrophes) unless it ships its own quoting. Without a key the placeholder &lt;key&gt; is shown.
    /// </summary>
    public static string BindLine(string resolvedCommand, string? bindKey, bool rawBind)
    {
        string key = string.IsNullOrWhiteSpace(bindKey) ? "<key>" : bindKey;
        if (rawBind) return $"bind {key} {resolvedCommand}";
        bool needsQuotes = resolvedCommand.Contains(';') || resolvedCommand.Contains(' ');
        string body = needsQuotes && !resolvedCommand.StartsWith('"') ? $"\"{resolvedCommand.Replace("\"", "'")}\"" : resolvedCommand;
        return $"bind {key} {body}";
    }

    /// <summary>Reference search: title, description, command or group, ignoring case.</summary>
    public static bool Matches(ConsoleCommandSpec command, string? query, string title, string description)
    {
        string text = (query ?? "").Trim();
        return text.Length == 0
            || title.Contains(text, StringComparison.OrdinalIgnoreCase)
            || description.Contains(text, StringComparison.OrdinalIgnoreCase)
            || command.Command.Contains(text, StringComparison.OrdinalIgnoreCase)
            || command.Group.Contains(text, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Reference Grouped: non-featured commands by group ("Other" when empty), groups and titles alphabetical.</summary>
    public static IReadOnlyList<IGrouping<string, T>> Grouped<T>(IEnumerable<T> commands, Func<T, ConsoleCommandSpec> spec, Func<T, string> title) =>
        commands.Where(command => !spec(command).Featured)
            .OrderBy(title, StringComparer.OrdinalIgnoreCase)
            .GroupBy(command => string.IsNullOrWhiteSpace(spec(command).Group) ? "Other" : spec(command).Group)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase).ToList();
}
