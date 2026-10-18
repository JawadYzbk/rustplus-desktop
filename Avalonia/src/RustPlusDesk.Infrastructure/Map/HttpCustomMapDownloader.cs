using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Map;

/// <summary>Reference DownloadCustomImageFromUrlAsync: a 20 s GET with the reference browser User-Agent; null on failure or no bytes.</summary>
public sealed class HttpCustomMapDownloader(HttpMessageHandler? handler = null) : ICustomMapDownloader
{
    public async Task<byte[]?> DownloadAsync(string url, CancellationToken ct)
    {
        try
        {
            using var client = handler is null ? new HttpClient() : new HttpClient(handler, disposeHandler: false);
            client.Timeout = TimeSpan.FromSeconds(20);
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
            byte[] bytes = await client.GetByteArrayAsync(url, ct).ConfigureAwait(false);
            return bytes.Length == 0 ? null : bytes;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or InvalidOperationException or UriFormatException) { return null; }
    }
}
