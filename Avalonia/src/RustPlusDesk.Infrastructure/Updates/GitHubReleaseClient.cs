using System.Text.Json;
using RustPlusDesk.Core.Updates;

namespace RustPlusDesk.Infrastructure.Updates;

/// <summary>
/// Reference UpdateService.GetLatestReleaseFromGitHubAsync: the latest published release is read from the
/// release API, and the installer asset is what the user is offered.
///
/// The preview asks its own repository, because the release that replaces the preview is the preview's release and
/// not the Windows app's.
/// </summary>
public sealed class GitHubReleaseClient(HttpClient http, string owner, string repository)
{
    public static string DefaultOwner = "JawadYzbk";
    public static string DefaultRepository = "rustplus-desktop";

    /// <summary>The page with the releases, opened when the user wants the new version.</summary>
    public string ReleasesPageUrl => $"https://github.com/{owner}/{repository}/releases/latest";

    /// <summary>
    /// The published releases, newest first, with the notes that go with them.
    ///
    /// The old app's patch notes window shows what each release changed; this reads the same list from the release
    /// API, so the notes are the ones the release actually has rather than a frozen copy of them.
    /// </summary>
    public async Task<IReadOnlyList<ReleaseNote>> GetReleasesAsync(int count = 20, CancellationToken ct = default)
    {
        var notes = new List<ReleaseNote>();
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get,
                $"https://api.github.com/repos/{owner}/{repository}/releases?per_page={Math.Clamp(count, 1, 50)}");
            request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("RustPlusDesk", AppVersion.Short(AppVersion.Raw()).ToString()));
            request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");

            using var response = await http.SendAsync(request, ct).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode) return notes;

            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            if (document.RootElement.ValueKind != JsonValueKind.Array) return notes;

            foreach (var release in document.RootElement.EnumerateArray())
            {
                string tag = release.TryGetProperty("tag_name", out var tagValue) ? tagValue.GetString() ?? "" : "";
                if (!Version.TryParse(AppVersion.Normalize(tag), out Version? version)) continue;

                string title = release.TryGetProperty("name", out var nameValue) ? nameValue.GetString() ?? tag : tag;
                string body = release.TryGetProperty("body", out var bodyValue) ? bodyValue.GetString() ?? "" : "";
                string url = release.TryGetProperty("html_url", out var urlValue) ? urlValue.GetString() ?? "" : "";
                bool preRelease = release.TryGetProperty("prerelease", out var pre) && pre.ValueKind == JsonValueKind.True;
                DateTime? published = release.TryGetProperty("published_at", out var publishedValue)
                    && publishedValue.TryGetDateTime(out DateTime moment) ? moment : null;

                notes.Add(new ReleaseNote(version, title, body, url, published, preRelease));
            }
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException or IOException)
        {
            // No releases is an empty list, not a broken app.
        }

        return notes;
    }

    /// <summary>The newest release, or null when the repository has none or the network is not there.</summary>
    public async Task<ReleaseInfo?> GetLatestAsync(string installerAssetName, CancellationToken ct)
    {
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get,
                $"https://api.github.com/repos/{owner}/{repository}/releases/latest");
            request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("RustPlusDesk", AppVersion.Short(AppVersion.Raw()).ToString()));
            request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");

            using var response = await http.SendAsync(request, ct).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode) return null;

            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            var root = document.RootElement;

            string tag = root.TryGetProperty("tag_name", out var tagValue) ? tagValue.GetString() ?? "" : "";
            string name = root.TryGetProperty("name", out var nameValue) ? nameValue.GetString() ?? tag : tag;
            if (!Version.TryParse(AppVersion.Normalize(tag), out Version? version)) return null;

            string? download = null;
            long? size = null;
            if (root.TryGetProperty("assets", out var assets) && assets.ValueKind == JsonValueKind.Array)
            {
                foreach (var asset in assets.EnumerateArray())
                {
                    string assetName = asset.TryGetProperty("name", out var value) ? value.GetString() ?? "" : "";
                    if (!string.Equals(assetName, installerAssetName, StringComparison.OrdinalIgnoreCase)) continue;
                    download = asset.TryGetProperty("browser_download_url", out var url) ? url.GetString() : null;
                    size = asset.TryGetProperty("size", out var sizeValue) && sizeValue.TryGetInt64(out long bytes) ? bytes : null;
                    break;
                }
            }

            return new ReleaseInfo(version, tag, name, download, size) { PageUrl = ReleasesPageUrl };
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException or IOException)
        {
            return null;
        }
    }
}
