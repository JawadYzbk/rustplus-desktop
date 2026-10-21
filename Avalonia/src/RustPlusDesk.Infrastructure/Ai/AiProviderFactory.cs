using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference AiProviderFactory: the provider the settings name.
///
/// The reference's fallback is OpenAI for any name it does not recognize, which keeps an older settings file working
/// when a provider is removed rather than failing every question.
/// </summary>
public static class AiProviderFactory
{
    public static IAiProvider For(string provider, HttpClient http, OpenRouterSettings? openRouter = null) => provider switch
    {
        AiProviders.Anthropic => new AnthropicProvider(http),
        AiProviders.Gemini => new GeminiProvider(http),
        AiProviders.OpenRouter => new OpenRouterProvider(http, openRouter),
        _ => new OpenAiProvider(http)
    };
}
