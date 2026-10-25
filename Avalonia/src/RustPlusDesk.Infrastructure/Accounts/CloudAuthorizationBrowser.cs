using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using RustPlusDesk.Core.Accounts;

namespace RustPlusDesk.Infrastructure.Accounts;

public interface ICloudAuthorizationBrowser
{
    Task<string> GetCodeAsync(CancellationToken cancellationToken);
}

/// <summary>Owns a system-browser loopback handoff, including cancellation and timeout cleanup.</summary>
public sealed class CloudAuthorizationBrowser(Action<Uri>? openBrowser = null) : ICloudAuthorizationBrowser
{
    public async Task<string> GetCodeAsync(CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromMinutes(3));
        var ct = timeout.Token;
        ct.ThrowIfCancellationRequested();
        using var listener = new HttpListener();
        // Reserve an ephemeral port briefly; retry if another process wins the bind race.
        Uri callback = null!;
        for (int attempt = 0; ; attempt++)
        {
            using var reservation = new TcpListener(IPAddress.Loopback, 0);
            reservation.Start();
            int port = ((IPEndPoint)reservation.LocalEndpoint).Port;
            reservation.Stop();
            callback = new Uri($"http://127.0.0.1:{port}/cloud/{Convert.ToHexString(RandomNumberGenerator.GetBytes(32))}/");
            listener.Prefixes.Clear();
            listener.Prefixes.Add(callback.AbsoluteUri);
            try { listener.Start(); break; }
            catch (HttpListenerException) when (attempt < 4) { }
        }
        using var cancel = ct.Register(listener.Close);
        try
        {
            var start = new Uri("https://rustplusdesktop.cloud/desktop/connect?redirect_uri=" + Uri.EscapeDataString(callback.AbsoluteUri));
            if (openBrowser is not null) openBrowser(start);
            else Process.Start(new ProcessStartInfo(start.AbsoluteUri) { UseShellExecute = true });
            while (true)
            {
                var context = await listener.GetContextAsync().ConfigureAwait(false);
                ct.ThrowIfCancellationRequested();
                var request = context.Request;
                string? code = request.QueryString["code"];
                bool valid = request.HttpMethod == "GET" && request.RemoteEndPoint is { } remote && IPAddress.IsLoopback(remote.Address)
                    && request.Url is { } url && url.Scheme == callback.Scheme && url.Authority == callback.Authority
                    && url.AbsolutePath == callback.AbsolutePath && url.Fragment.Length == 0
                    && request.QueryString.Count == 1 && request.QueryString.GetValues("code")?.Length == 1
                    && code is { Length: 64 } && code.All(char.IsAsciiLetterOrDigit);
                context.Response.StatusCode = valid ? 200 : 400;
                context.Response.ContentType = "text/plain; charset=utf-8";
                context.Response.Headers[HttpResponseHeader.CacheControl] = "no-store";
                context.Response.Headers["Referrer-Policy"] = "no-referrer";
                byte[] body = System.Text.Encoding.UTF8.GetBytes(valid
                    ? "Authorization received. Return to Rust+ Desktop to check sign-in."
                    : "Invalid sign-in callback. Return to Rust+ Desktop.");
                context.Response.ContentLength64 = body.Length;
                try { await context.Response.OutputStream.WriteAsync(body, ct).ConfigureAwait(false); }
                catch (IOException) { /* A closed tab must not consume a different callback. */ }
                finally { context.Response.Close(); }
                if (valid) return code!;
            }
        }
        catch (Exception) when (ct.IsCancellationRequested) { throw new OperationCanceledException(ct); }
        catch (Exception) { throw new CloudAccountException(CloudAccountFailure.Unavailable); }
        finally { listener.Close(); }
    }
}
