using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference AiProviderFactory: the provider the settings name.
///
/// OpenAI and OpenRouter speak the same chat-completions shape, so one implementation serves both; the reference also
/// carries Anthropic and Gemini, which use their own shapes and are not ported yet.
/// </summary>
public static class AiProviderFactory
{
    public static IAiProvider For(string provider, HttpClient http) => provider switch
    {
        AiProviders.OpenRouter => new OpenAiProvider(http, "https://openrouter.ai/api/v1"),
        _ => new OpenAiProvider(http)
    };
}
