namespace RustPlusDesk.Core.Ai;

/// <summary>
/// Everything the AI companion needs, stored on this machine and nowhere else.
///
/// The key never reaches the cloud: the settings file is written with the current user's permissions only, so another
/// account on the same machine cannot read it.
/// </summary>
public sealed record AiSettings
{
    /// <summary>Which provider answers (reference AiProviders).</summary>
    public string Provider { get; init; } = AiProviders.OpenAi;

    /// <summary>The model the provider is asked for.</summary>
    public string Model { get; init; } = "gpt-4o-mini";

    /// <summary>The provider's key. It is written to a private file and never sent anywhere but the provider.</summary>
    public string ApiKey { get; init; } = "";

    /// <summary>Whether the companion is set up at all (reference AiCompanionStore.HasKey).</summary>
    public bool IsSetUp => ApiKey.Trim().Length > 0;
}

/// <summary>Reads and writes the AI companion's settings.</summary>
public interface IAiSettingsStore
{
    Task<AiSettings> LoadAsync(CancellationToken ct);
    Task SaveAsync(AiSettings settings, CancellationToken ct);
}
