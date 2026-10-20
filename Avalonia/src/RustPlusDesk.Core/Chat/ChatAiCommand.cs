using System.Text;

namespace RustPlusDesk.Core.Chat;

/// <summary>
/// Reference the AI chat command: ask the companion from inside the game, and get the answer in chat.
///
/// It is the one command that costs money to answer, and the money is the owner's. That shapes everything here: it
/// is off until it is set up, it is capped per person per hour once it is, and it never spends anything before both
/// of those have been checked.
/// </summary>
public static class ChatAiCommand
{
    /// <summary>
    /// Rust truncates a long chat line, so the answer is asked for short and then cut anyway. Twenty-five words is
    /// about one to two concise sentences, matching the in-game chat limits.
    /// </summary>
    public const int MaxWords = 25;

    /// <summary>The hard ceiling in characters, applied after the model has had its say.</summary>
    public const int MaxChars = 175;

    /// <summary>
    /// Unlike every other command this one takes an argument, so it matches on the first word rather than on the
    /// whole line.
    /// </summary>
    public static bool Matches(string command, string aiCommand)
    {
        string word = aiCommand.Trim().ToLowerInvariant();
        if (word.Length == 0) return false;
        return command == word || command.StartsWith(word + " ", StringComparison.Ordinal);
    }

    /// <summary>
    /// The question, taken from the raw message rather than the lower-cased copy the dispatcher matches on: names,
    /// monuments and the question mark all matter to the model.
    /// </summary>
    public static string Question(string rawText, string prefix, string aiCommand)
    {
        string text = rawText.Trim();
        int cut = Math.Min(text.Length, prefix.Length + aiCommand.Trim().Length);
        return text[cut..].Trim();
    }

    /// <summary>
    /// Cuts the answer to something the game will actually show, at a sentence end where there is one, because a reply
    /// that stops mid-clause reads as a broken app rather than a long answer.
    /// </summary>
    public static string Shorten(string answer)
    {
        answer = Clean(answer);
        if (answer.Length <= MaxChars) return answer;

        var cut = answer[..MaxChars];
        int stop = cut.LastIndexOfAny(['.', '!', '?']);
        return stop > MaxChars / 2 ? cut[..(stop + 1)] : cut.TrimEnd() + "…";
    }

    /// <summary>Reference AiPrompt.CleanResponse: the model likes to wrap its answer, and that is not the answer.</summary>
    public static string Clean(string answer)
    {
        string text = answer.Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            int firstBreak = text.IndexOf('\n');
            int lastBreak = text.LastIndexOf("```", StringComparison.Ordinal);
            if (firstBreak > 0 && lastBreak > firstBreak) text = text[(firstBreak + 1)..lastBreak].Trim();
        }
        return text.Trim('"', ' ', '\n', '\r');
    }

    /// <summary>
    /// Whether this author would be allowed to use it, for the command list. The allowance is deliberately not
    /// checked here: someone who has used their five for the hour still has the command.
    /// </summary>
    public static bool IsListable(bool isSetUp, bool isPremium, bool isOwner, bool allowTeammates) =>
        isSetUp && isPremium && (isOwner || allowTeammates);

    /// <summary>
    /// Whether this author has an ask left in the current hour. A rolling hour rather than a clock hour, so the
    /// allowance cannot be doubled by asking twice either side of the top of the hour.
    /// </summary>
    public static bool HasAllowanceLeft(int limit, IReadOnlyList<DateTime> asks, DateTime now)
    {
        if (limit <= 0) return true;
        int used = asks.Count(at => at >= now.AddHours(-1));
        return used < limit;
    }
}
