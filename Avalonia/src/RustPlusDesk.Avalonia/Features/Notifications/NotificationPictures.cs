using System.Collections.Concurrent;
using System.Net.Http;
using global::Avalonia.Media.Imaging;

namespace RustPlusDesk.Avalonia.Features.Notifications;

/// <summary>
/// The server's picture, as the push sent it.
///
/// The reference's notification centre shows the picture that came with the push - the server's logo or banner - and
/// falls back to the type's glyph when there is none. The picture is fetched once and kept, because the same server's
/// picture arrives with every one of its notifications.
/// </summary>
public static class NotificationPictures
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(10) };
    private static readonly ConcurrentDictionary<string, Task<Bitmap?>> Cache = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The picture at a URL, or null when there is none or it cannot be fetched.
    ///
    /// A picture that cannot be read is no picture: the row shows its glyph instead of a broken frame.
    /// </summary>
    public static Task<Bitmap?> LoadAsync(string? url)
    {
        if (string.IsNullOrWhiteSpace(url)) return Task.FromResult<Bitmap?>(null);
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) return Task.FromResult<Bitmap?>(null);
        if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps) return Task.FromResult<Bitmap?>(null);

        return Cache.GetOrAdd(url, key => FetchAsync(key));
    }

    private static async Task<Bitmap?> FetchAsync(string url)
    {
        try
        {
            byte[] bytes = await Http.GetByteArrayAsync(url).ConfigureAwait(false);
            if (bytes.Length == 0) return null;

            using var stream = new MemoryStream(bytes, writable: false);
            return new Bitmap(stream);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or IOException or NotSupportedException)
        {
            return null;
        }
    }
}
