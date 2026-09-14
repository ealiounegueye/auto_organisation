using System.Windows;
using System.Windows.Interop;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using OutlookOrganizer.Configuration;

namespace OutlookOrganizer.Services;

public sealed class GraphAuthService
{
    public static readonly string[] Scopes =
    [
        "User.Read",
        "Mail.ReadWrite",
        "MailboxSettings.ReadWrite"
    ];

    private readonly IPublicClientApplication _app;
    private readonly Func<IntPtr> _getWindowHandle;
    private bool _cacheReady;

    public GraphAuthService(AppOptions options, Func<IntPtr>? getWindowHandle = null)
    {
        if (!options.HasValidClientId)
        {
            throw new InvalidOperationException(
                "Le Client ID Microsoft Entra n'est pas configuré. Copiez appsettings.local.json.example vers appsettings.local.json et renseignez l'identifiant d'application.");
        }

        _getWindowHandle = getWindowHandle ?? (() => IntPtr.Zero);

        var builder = PublicClientApplicationBuilder
            .Create(options.AzureAd.ClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, options.AzureAd.TenantId)
            .WithRedirectUri(string.IsNullOrWhiteSpace(options.AzureAd.RedirectUri)
                ? "http://localhost"
                : options.AzureAd.RedirectUri)
            .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows)
            {
                Title = "Outlook Organizer"
            });

        _app = builder.Build();
    }

    public async Task InitializeCacheAsync()
    {
        if (_cacheReady)
        {
            return;
        }

        await TokenCacheHelper.RegisterAsync(_app);
        _cacheReady = true;
    }

    public async Task<AuthenticationResult> SignInAsync(CancellationToken cancellationToken = default)
    {
        var accounts = await _app.GetAccountsAsync();
        var account = accounts.FirstOrDefault();

        try
        {
            return await _app
                .AcquireTokenSilent(Scopes, account)
                .ExecuteAsync(cancellationToken);
        }
        catch (MsalUiRequiredException)
        {
            return await AcquireInteractiveAsync(cancellationToken);
        }
    }

    public async Task<AuthenticationResult?> TrySignInSilentlyAsync(CancellationToken cancellationToken = default)
    {
        var accounts = await _app.GetAccountsAsync();
        var account = accounts.FirstOrDefault();
        if (account is null)
        {
            return null;
        }

        try
        {
            return await _app
                .AcquireTokenSilent(Scopes, account)
                .ExecuteAsync(cancellationToken);
        }
        catch (MsalUiRequiredException)
        {
            return null;
        }
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default)
    {
        var result = await SignInAsync(cancellationToken);
        return result.AccessToken;
    }

    public async Task SignOutAsync()
    {
        var accounts = await _app.GetAccountsAsync();
        foreach (var account in accounts)
        {
            await _app.RemoveAsync(account);
        }
    }

    private async Task<AuthenticationResult> AcquireInteractiveAsync(CancellationToken cancellationToken)
    {
        var interactive = _app
            .AcquireTokenInteractive(Scopes)
            .WithPrompt(Prompt.SelectAccount)
            .WithParentActivityOrWindow(_getWindowHandle);

        try
        {
            return await interactive.ExecuteAsync(cancellationToken);
        }
        catch (MsalClientException)
        {
            return await _app
                .AcquireTokenInteractive(Scopes)
                .WithPrompt(Prompt.SelectAccount)
                .WithUseEmbeddedWebView(false)
                .ExecuteAsync(cancellationToken);
        }
    }

    public static IntPtr GetActiveWindowHandle()
    {
        if (Application.Current?.MainWindow is { } window)
        {
            return new WindowInteropHelper(window).Handle;
        }

        return IntPtr.Zero;
    }
}
