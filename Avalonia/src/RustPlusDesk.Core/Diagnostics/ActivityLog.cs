using System.Text.RegularExpressions;

namespace RustPlusDesk.Core.Diagnostics;

/// <summary>
/// The reference diagnostics log (MainWindow.AppendLog): "[HH:mm:ss.fff] line", newest last, the oldest lines
/// dropped beyond <see cref="MaxLines"/>, filtered by a case-insensitive substring. Lines are redacted before
/// they are stored so a token, password or JWT never reaches the screen or a screenshot.
/// </summary>
public sealed class ActivityLog(TimeProvider? clock = null)
{
    public const int MaxLines = 2000;
    private readonly TimeProvider clock = clock ?? TimeProvider.System;
    private readonly List<string> lines = [];

    public IReadOnlyList<string> Lines => lines;

    /// <summary>Adds a timestamped, redacted line; returns false for blank input.</summary>
    public bool Append(string? line)
    {
        if (string.IsNullOrWhiteSpace(line)) return false;
        lines.Add($"[{clock.GetLocalNow():HH:mm:ss.fff}] {LogRedactor.Redact(line.Trim())}");
        if (lines.Count > MaxLines) lines.RemoveRange(0, lines.Count - MaxLines);
        return true;
    }

    public IEnumerable<string> Filter(string? query)
    {
        string trimmed = query?.Trim() ?? "";
        return trimmed.Length == 0 ? lines : lines.Where(line => line.Contains(trimmed, StringComparison.OrdinalIgnoreCase));
    }

    public void Clear() => lines.Clear();
}

/// <summary>Masks credentials that status or error text could carry (key=value secrets, bearer tokens, JWTs).</summary>
public static partial class LogRedactor
{
    public const string Mask = "***";

    public static string Redact(string line)
    {
        line = SecretAssignment().Replace(line, match => match.Groups["key"].Value + match.Groups["sep"].Value + Mask);
        line = Bearer().Replace(line, "Bearer " + Mask);
        return Jwt().Replace(line, Mask);
    }

    [GeneratedRegex(@"(?<key>\b[\w-]*(?:token|password|passwd|secret|api[_-]?key|auth[_-]?key|credential|session[_-]?id)\b""?)(?<sep>\s*[:=]\s*""?)[^\s"",;&]+",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex SecretAssignment();

    [GeneratedRegex(@"\bBearer\s+[A-Za-z0-9._~+/=-]+", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex Bearer();

    [GeneratedRegex(@"\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]+", RegexOptions.CultureInvariant)]
    private static partial Regex Jwt();
}
