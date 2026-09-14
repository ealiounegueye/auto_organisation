using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace OutlookOrganizer.Services;

public static class TokenCacheHelper
{
    public static async Task RegisterAsync(IPublicClientApplication app)
    {
        var folder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookOrganizer");

        Directory.CreateDirectory(folder);

        var storage = new StorageCreationPropertiesBuilder("msalcache.bin", folder).Build();
        var cacheHelper = await MsalCacheHelper.CreateAsync(storage);
        cacheHelper.RegisterCache(app.UserTokenCache);
    }
}
