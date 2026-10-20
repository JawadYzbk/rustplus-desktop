using System.Text.Json.Nodes;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile chat-command fields in profiles.json (ChatCommandsEnabled, ChatCommandPrefix,
/// ChatResponseDelaySeconds and the Cmd* names). Changes go through the serialized, backed-up profile writer and
/// adopt the latest document; only this profile's fields are replaced.
/// </summary>
public sealed class LegacyChatCommandStore : IChatCommandSettingsStore
{
    public async Task<ChatCommandSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        return bytes is null ? ChatCommandSettings.Defaults : Read(LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey));
    }

    public async Task SaveAsync(string profilesPath, string profileKey, ChatCommandSettings settings, CancellationToken ct) =>
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node["ChatCommandsEnabled"] = settings.Enabled;
            node["ChatCommandPrefix"] = settings.Prefix;
            node["ChatResponseDelaySeconds"] = settings.ResponseDelaySeconds;
            node["CmdPop"] = settings.Pop;
            node["CmdList"] = settings.List;
            node["CmdTime"] = settings.Time;
            node["CmdPromote"] = settings.Promote;
            node["CmdDeepSea"] = settings.DeepSea;
            node["CmdCargo"] = settings.Cargo;
            node["CmdAfk"] = settings.Afk;
            node["CmdOilRig"] = settings.OilRig;
            node["CmdHeli"] = settings.Heli;
            node["CmdVendor"] = settings.Vendor;
            node["CmdUpkeepDetail"] = settings.UpkeepDetail;
            node["CmdBaseCodes"] = settings.BaseCodes;
            node["CmdAi"] = settings.Ai;
            node["CmdCustomTimer"] = settings.CustomTimer;
            node["ClanCommandsAllowBaseCodes"] = settings.ClanCommandsAllowBaseCodes;
            node["ClanCommandsAllowSwitches"] = settings.ClanCommandsAllowSwitches;
            node["ClanChatCommandsEnabled"] = settings.ClanCommandsEnabled;
            node["ClanCommandsAllowAi"] = settings.ClanCommandsAllowAi;
            node["ClanRolePermissions"] = new JsonArray(settings.ClanRolePermissions.Select(role => (JsonNode)new JsonObject
            {
                ["RoleId"] = role.RoleId,
                ["Name"] = role.Name,
                ["Rank"] = role.Rank,
                ["Allowed"] = role.Allowed
            }).ToArray());
            node["BaseCodes"] = new JsonArray(settings.BaseCodeRows.Select(code => (JsonNode)new JsonObject
            {
                ["Label"] = code.Label,
                ["Code"] = code.Code
            }).ToArray());
        }, null, ct).ConfigureAwait(false);

    private static ChatCommandSettings Read(JsonObject? profile)
    {
        var defaults = ChatCommandSettings.Defaults;
        if (profile is null) return defaults;
        return new ChatCommandSettings(
            Enabled: Flag(profile, "ChatCommandsEnabled") ?? defaults.Enabled,
            Prefix: Prefix(profile, defaults.Prefix),
            ResponseDelaySeconds: Number(profile, "ChatResponseDelaySeconds") ?? defaults.ResponseDelaySeconds,
            Pop: Command(profile, "CmdPop", defaults.Pop),
            List: Command(profile, "CmdList", defaults.List),
            Time: Command(profile, "CmdTime", defaults.Time),
            Promote: Command(profile, "CmdPromote", defaults.Promote),
            DeepSea: Command(profile, "CmdDeepSea", defaults.DeepSea),
            Cargo: Command(profile, "CmdCargo", defaults.Cargo),
            Afk: Command(profile, "CmdAfk", defaults.Afk),
            OilRig: Command(profile, "CmdOilRig", defaults.OilRig),
            Heli: Command(profile, "CmdHeli", defaults.Heli),
            Vendor: Command(profile, "CmdVendor", defaults.Vendor),
            UpkeepDetail: Command(profile, "CmdUpkeepDetail", defaults.UpkeepDetail),
            BaseCodes: Command(profile, "CmdBaseCodes", defaults.BaseCodes),
            Ai: Command(profile, "CmdAi", defaults.Ai),
            CustomTimer: Command(profile, "CmdCustomTimer", defaults.CustomTimer))
        {
            SwitchMappings = Mappings(profile, "SwitchCommandMappings"),
            UpkeepMappings = Mappings(profile, "UpkeepCommandMappings"),
            BaseCodeRows = BaseCodes(profile),
            ClanCommandsAllowBaseCodes = Flag(profile, "ClanCommandsAllowBaseCodes") ?? false,
            ClanCommandsAllowSwitches = Flag(profile, "ClanCommandsAllowSwitches") ?? false,
            ClanCommandsEnabled = Flag(profile, "ClanChatCommandsEnabled") ?? false,
            ClanCommandsAllowAi = Flag(profile, "ClanCommandsAllowAi") ?? false,
            ClanRolePermissions = RolePermissions(profile)
        };
    }

    /// <summary>Reference ServerProfile.ClanRolePermissions rows: the role, its name and rank, and whether it may command.</summary>
    private static IReadOnlyList<ChatRolePermission> RolePermissions(JsonObject profile)
    {
        if (profile["ClanRolePermissions"] is not JsonArray rows) return ChatCommandSettings.NoRolePermissions;
        var permissions = new List<ChatRolePermission>();
        foreach (var row in rows)
        {
            if (row is not JsonObject role) continue;
            if (role["RoleId"] is not JsonValue id || !id.TryGetValue(out int roleId)) continue;
            permissions.Add(new ChatRolePermission(roleId, Text(role, "Name") ?? "", (int)(Number(role, "Rank") ?? 0),
                Flag(role, "Allowed") ?? false));
        }
        return permissions.Count == 0 ? ChatCommandSettings.NoRolePermissions : permissions;
    }

    /// <summary>Reference ServerProfile.BaseCodes rows; blank rows are spare editor slots and are not carried over.</summary>
    private static IReadOnlyList<ChatBaseCode> BaseCodes(JsonObject profile)
    {
        if (profile["BaseCodes"] is not JsonArray rows) return ChatCommandSettings.NoBaseCodes;
        var codes = new List<ChatBaseCode>();
        foreach (var row in rows)
        {
            if (row is not JsonObject code) continue;
            string label = Text(code, "Label") ?? "";
            string value = Text(code, "Code") ?? "";
            if (string.IsNullOrWhiteSpace(label) && string.IsNullOrWhiteSpace(value)) continue;
            codes.Add(new ChatBaseCode(label, value));
        }
        return codes.Count == 0 ? ChatCommandSettings.NoBaseCodes : codes;
    }

    /// <summary>Reference ChatCommandMapping rows: the entity the command is bound to and the command name.</summary>
    private static IReadOnlyList<ChatCommandMapping> Mappings(JsonObject profile, string name)
    {
        if (profile[name] is not JsonArray rows) return ChatCommandSettings.NoMappings;
        var mappings = new List<ChatCommandMapping>();
        foreach (var row in rows)
        {
            if (row is not JsonObject mapping) continue;
            uint entityId = mapping["EntityId"] is JsonValue id && id.TryGetValue(out uint value) ? value : 0;
            if (entityId == 0) continue;
            mappings.Add(new ChatCommandMapping(entityId, Text(mapping, "Command") ?? ""));
        }
        return mappings.Count == 0 ? ChatCommandSettings.NoMappings : mappings;
    }

    private static bool? Flag(JsonObject profile, string name) =>
        profile[name] is JsonValue value && value.TryGetValue(out bool flag) ? flag : null;

    private static string? Text(JsonObject profile, string name) =>
        profile[name] is JsonValue value && value.TryGetValue(out string? text) ? text : null;

    private static double? Number(JsonObject profile, string name) =>
        profile[name] is JsonValue value && value.TryGetValue(out double number) ? number : null;

    private static string Command(JsonObject profile, string name, string fallback) =>
        ChatCommandSettings.ValidateCommand(Text(profile, name), fallback);

    /// <summary>Reference ChatCommandPrefix setter: only the four accepted prefixes are stored; anything else reads as "!".</summary>
    private static string Prefix(JsonObject profile, string fallback)
    {
        string? text = Text(profile, "ChatCommandPrefix");
        return ChatCommandSettings.IsValidPrefix(text) ? text! : fallback;
    }
}
