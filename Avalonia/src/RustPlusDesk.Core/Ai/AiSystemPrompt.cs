namespace RustPlusDesk.Core.Ai;

/// <summary>
/// The companion's standing instructions (reference AiPrompt.SystemMessage).
///
/// The reference sends the raid table and every recipe with the question. The port asks for the same answer without
/// carrying the tables, so the companion has to be asked for what it does not know - and says so instead of guessing.
/// </summary>
public static class AiSystemPrompt
{
    public static string For(AiQuestion question)
    {
        if (!string.IsNullOrWhiteSpace(question.Instructions)) return question.Instructions!;

        var prompt = new System.Text.StringBuilder();
        prompt.Append("You are a companion for the game Rust. You answer in one or two short sentences, ")
            .Append("the way a teammate would answer in chat. Say plainly when you do not know something ")
            .Append("rather than making it up, and never invent item names, prices or numbers.");
        if (!string.IsNullOrWhiteSpace(question.Context))
            prompt.Append("\n\nWhat is happening right now: ").Append(question.Context);
        if (question.MaxWords > 0) prompt.Append("\n\nAnswer in at most ").Append(question.MaxWords).Append(" words.");
        if (!question.MatchQuestionLanguage && !string.IsNullOrWhiteSpace(question.Language))
            prompt.Append("\n\nAnswer in ").Append(question.Language).Append('.');
        return prompt.ToString();
    }
}
