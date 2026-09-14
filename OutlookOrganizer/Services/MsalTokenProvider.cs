using Microsoft.Kiota.Abstractions.Authentication;

namespace OutlookOrganizer.Services;

public sealed class MsalTokenProvider : IAccessTokenProvider
{
    private readonly GraphAuthService _auth;

    public MsalTokenProvider(GraphAuthService auth)
    {
        _auth = auth;
        AllowedHostsValidator = new AllowedHostsValidator(["graph.microsoft.com"]);
    }

    public AllowedHostsValidator AllowedHostsValidator { get; }

    public async Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
    {
        return await _auth.GetAccessTokenAsync(cancellationToken);
    }
}
