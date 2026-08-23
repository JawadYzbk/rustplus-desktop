/**
 * Settings search matcher — port of Views/MainWindow/Map/SettingsSearchMatcher.cs.
 * Every whitespace-separated query term must appear (case-insensitive) in at least one
 * of the provided values (title or keywords).
 */
export function settingsSearchMatches(query: string, ...values: string[]): boolean {
  const terms = query
    .split(" ")
    .map((t) => t.trim())
    .filter((t) => t.length > 0);
  const haystacks = values.map((v) => v.toLowerCase());
  return terms.every((term) => haystacks.some((value) => value.includes(term.toLowerCase())));
}
