using System.Net;
using System.Net.Http.Headers;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

[PublicAPI]
public abstract class OAuthAccessTokenCredentials : ExchangeCredentials
{
    private const string BearerAuthenticationType = "Bearer";

    /// <summary>
    /// Handler for acquiring the access token for each ews request
    /// </summary>
    /// <returns></returns>
    public abstract Task<string> AcquireAccessToken();


    internal override async System.Threading.Tasks.Task PrepareWebRequest(EwsHttpWebRequest request)
    {
        var token = await AcquireAccessToken().ConfigureAwait(false);

        request.Headers.Remove(HttpRequestHeader.Authorization.ToString());
        request.Headers.Authorization = new AuthenticationHeaderValue(BearerAuthenticationType, token);
    }
}
