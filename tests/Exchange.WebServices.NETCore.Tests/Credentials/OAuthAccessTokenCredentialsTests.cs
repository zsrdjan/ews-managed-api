using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web.TokenCacheProviders;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Credentials;

internal class TokenProvider : OAuthAccessTokenCredentials
{
    private static readonly string[] EwsScopes = ["https://outlook.office365.com/.default",];

    private readonly IConfidentialClientApplication _cca;


    public TokenProvider(OutlookConnectionOptions options, IMsalTokenCacheProvider cacheProvider)
    {
        var applicationOptions = new ConfidentialClientApplicationOptions
        {
            ClientId = options.ClientId,
            ClientSecret = options.ClientSecret,
            TenantId = options.TenantId,
        };

        _cca = ConfidentialClientApplicationBuilder.CreateWithApplicationOptions(applicationOptions).Build();

        // Attach cache provider
        cacheProvider.Initialize(_cca.UserTokenCache);
        cacheProvider.Initialize(_cca.AppTokenCache);
    }

    public override async Task<string> AcquireAccessToken()
    {
        var result = await _cca.AcquireTokenForClient(EwsScopes).ExecuteAsync();

        return result.AccessToken;
    }
}

public class OAuthAccessTokenCredentialsTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public OAuthAccessTokenCredentialsTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task AuthenticationTest()
    {
        var options = _provider.OutlookConnectionOptions;

        var service = new ExchangeService
        {
            Credentials = new TokenProvider(options, _provider.GetRequiredService<IMsalTokenCacheProvider>()),
            UseDefaultCredentials = false,
            AcceptGzipEncoding = true,
            Url = new Uri(options.Url),
            ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, options.ImpersonationUpn),
        };

        var folders = await service.FindFolders(
            new FolderId(WellKnownFolderName.MsgFolderRoot),
            new FolderView(100, 0)
        );

        Assert.NotEmpty(folders);
    }
}
