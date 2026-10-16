namespace RustPlusDesk.Core.Accounts;

/// <summary>The desktop browser boundary; callback URLs are consumed by infrastructure, never view models.</summary>
public interface IAuthenticationBrowser
{
    Task<Uri?> AuthenticateAsync(Uri requestUri, Uri callbackUri, CancellationToken cancellationToken);
}
