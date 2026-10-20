using System.Diagnostics;
using RustPlusDesk.Core.Diagnostics;

namespace RustPlusDesk.Infrastructure.Diagnostics;

/// <summary>
/// Reference TrafficTrackingHttpMessageHandler: every request the app makes is measured and written down, so the
/// traffic window can show what went where. The handler never changes the request or the answer.
/// </summary>
public sealed class TrafficTrackingHttpMessageHandler(HttpMessageHandler innerHandler, NetworkTrafficMonitor monitor,
    string? category = null) : DelegatingHandler(innerHandler)
{
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        if (!monitor.IsEnabled) return await base.SendAsync(request, cancellationToken).ConfigureAwait(false);

        var stopwatch = Stopwatch.StartNew();
        string effectiveCategory = category ?? DetectCategoryFromUri(request.RequestUri);
        string endpoint = FormatEndpoint(effectiveCategory, $"{request.Method} {request.RequestUri?.ToString() ?? "HTTP"}");
        long bytesOut = 0;
        if (request.Content is not null)
        {
            try
            {
                if (request.Content.Headers.ContentLength.HasValue) bytesOut = request.Content.Headers.ContentLength.Value;
            }
            catch (Exception) { }
        }

        try
        {
            var response = await base.SendAsync(request, cancellationToken).ConfigureAwait(false);
            stopwatch.Stop();
            long bytesIn = 0;
            if (response.Content is not null && response.Content.Headers.ContentLength.HasValue)
                bytesIn = response.Content.Headers.ContentLength.Value;

            monitor.Record(effectiveCategory, TrafficDirection.Both, endpoint, bytesIn, bytesOut,
                stopwatch.ElapsedMilliseconds, $"{(int)response.StatusCode} {response.StatusCode}");
            return response;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            monitor.Record(effectiveCategory, TrafficDirection.Outbound, endpoint, 0, bytesOut,
                stopwatch.ElapsedMilliseconds, "Error", ex.Message);
            throw;
        }
    }

    /// <summary>Reference NetworkTrafficExtensions.DetectCategoryFromUri: the host says what the request was for.</summary>
    public static string DetectCategoryFromUri(Uri? uri)
    {
        if (uri is null || !uri.IsAbsoluteUri) return "HTTP";
        string host = uri.Host.ToLowerInvariant();
        if (host.Contains("supabase.co")) return "Cloud API";
        if (host.Contains("rustplusdesktop.cloud") || host.Contains("rustplus-desktop.fly.dev")) return "Cloud API";
        if (host.Contains("steamcommunity.com") || host.Contains("steampowered.com")) return "Steam";
        if (host.Contains("rusthelp.com")) return "Items";
        if (host.Contains("api.rustplus.io")) return "Rust+";
        return "HTTP";
    }

    /// <summary>
    /// Reference NetworkTrafficExtensions.FormatEndpoint: a cloud endpoint is shown as its path, so the log does not
    /// repeat a long host name on every line.
    /// </summary>
    public static string FormatEndpoint(string category, string endpoint)
    {
        if (string.IsNullOrWhiteSpace(endpoint)) return string.Empty;
        string trimmed = endpoint.Trim();
        if (!IsCloudCategory(category)) return trimmed;

        int spaceIndex = trimmed.IndexOf(' ');
        if (spaceIndex > 0)
        {
            string method = trimmed[..spaceIndex].Trim();
            string urlPart = trimmed[(spaceIndex + 1)..].Trim();
            if (Uri.TryCreate(urlPart, UriKind.Absolute, out var parsed)) return $"{method} {parsed.PathAndQuery}";
            return trimmed;
        }

        if (Uri.TryCreate(trimmed, UriKind.Absolute, out var direct)) return direct.PathAndQuery;
        return trimmed;
    }

    private static bool IsCloudCategory(string category) =>
        category.Contains("Cloud", StringComparison.OrdinalIgnoreCase);
}
