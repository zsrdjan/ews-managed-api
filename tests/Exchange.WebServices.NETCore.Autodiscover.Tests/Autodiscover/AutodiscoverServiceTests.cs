using System.Diagnostics;

using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Autodiscover.Tests.Autodiscover;

public class AutodiscoverServiceTests : IClassFixture<AutodiscoverProvider>
{
    private readonly AutodiscoverProvider _provider;

    public AutodiscoverServiceTests(AutodiscoverProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task DiscoveryTest()
    {
        var options = _provider.ConnectionOptions.Value;

        var service = new AutodiscoverService
        {
            Credentials = new WebCredentials(options.UserName, options.Password),
        };


        var result = await service.GetUserSettings(
            options.UserName,
            UserSettingName.AlternateMailboxes,
            UserSettingName.ExternalEwsUrl
        );

        Debugger.Break();
    }
}
