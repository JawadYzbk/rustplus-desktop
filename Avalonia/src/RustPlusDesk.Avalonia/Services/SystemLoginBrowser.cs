using System.Net;
using System.Text;
using Avalonia.Controls;
using Avalonia.Threading;
using RustPlusDesk.Core.Accounts;

namespace RustPlusDesk.Avalonia.Services;

/// <summary>Owns a loopback callback listener until browser completion or cancellation.</summary>
public sealed class SystemLoginBrowser : IAuthenticationBrowser
{
    private readonly Func<Uri, CancellationToken, Task> launch;
    private readonly Action activate;

    public SystemLoginBrowser(Window owner)
    {
        launch = async (uri, ct) => await Dispatcher.UIThread.InvokeAsync(async () =>
        {
            ct.ThrowIfCancellationRequested();
            if (!await owner.Launcher.LaunchUriAsync(uri)) throw new InvalidOperationException("Default browser unavailable.");
        });
        activate = () => Dispatcher.UIThread.Post(() =>
        {
            if (!owner.IsVisible) return;
            if (owner.WindowState == WindowState.Minimized) owner.WindowState = WindowState.Normal;
            owner.Activate();
        });
    }

    internal SystemLoginBrowser(Action<Uri> launch, Action activate)
    {
        this.launch = (uri, _) => { launch(uri); return Task.CompletedTask; };
        this.activate = activate;
    }

    public async Task<Uri?> AuthenticateAsync(Uri requestUri, Uri callbackUri, CancellationToken cancellationToken)
    {
        if (!requestUri.IsAbsoluteUri || !callbackUri.IsAbsoluteUri || requestUri.Scheme != Uri.UriSchemeHttps || requestUri.UserInfo.Length != 0 ||
            callbackUri.Scheme != Uri.UriSchemeHttp || callbackUri.Host != "127.0.0.1" || callbackUri.Port < 1024 ||
            callbackUri.UserInfo.Length != 0 || callbackUri.Query.Length != 0 || callbackUri.Fragment.Length != 0 ||
            !callbackUri.AbsolutePath.EndsWith('/')) throw new ArgumentException("Invalid authentication endpoint.");
        cancellationToken.ThrowIfCancellationRequested();
        using var listener = new HttpListener();
        listener.Prefixes.Add(callbackUri.AbsoluteUri);
        listener.Start();
        using var cancel = cancellationToken.Register(listener.Close);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            await launch(requestUri, cancellationToken).ConfigureAwait(false);
            while (true)
            {
                var context = await listener.GetContextAsync().ConfigureAwait(false);
                bool accepted = context.Request.RemoteEndPoint is { } remote && IPAddress.IsLoopback(remote.Address) &&
                    context.Request.HttpMethod == "GET" && context.Request.Url is { } url &&
                    url.GetLeftPart(UriPartial.Path) == callbackUri.AbsoluteUri;
                using var response = context.Response;
                response.StatusCode = accepted ? 200 : 400;
                response.ContentType = "text/html; charset=utf-8";
                response.Headers["Cache-Control"] = "no-store";
                response.Headers["Referrer-Policy"] = "no-referrer";
                response.Headers["Content-Security-Policy"] = "default-src 'none'";
                byte[] body = Encoding.UTF8.GetBytes("<!doctype html><meta charset='utf-8'><title>Rust+ Desktop</title><p>Return to Rust+ Desktop to check your sign-in. You can close this tab.</p>");
                response.ContentLength64 = body.Length;
                await response.OutputStream.WriteAsync(body, cancellationToken).ConfigureAwait(false);
                if (!accepted) continue;
                cancellationToken.ThrowIfCancellationRequested();
                activate();
                return context.Request.Url;
            }
        }
        catch (Exception ex) when (cancellationToken.IsCancellationRequested && ex is HttpListenerException or ObjectDisposedException or IOException)
        { throw new OperationCanceledException(cancellationToken); }
    }
}
