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
