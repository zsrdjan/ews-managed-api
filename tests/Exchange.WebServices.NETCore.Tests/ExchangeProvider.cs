using System.Reflection;

using JetBrains.Annotations;

using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace Exchange.WebServices.NETCore.Tests;

[PublicAPI]
public class ExchangeConnectionOptions
{
    public required string ServiceUrl { get; set; }

    public required string UserName { get; set; }

    public required string Password { get; set; }

    public required string ImpersonateUpn { get; set; }
}

[PublicAPI]
public class ExchangeProvider
{
    protected readonly IServiceProvider _provider = BuildServiceProvider();


    public IOptions<ExchangeConnectionOptions> ConnectionOptions =>
        _provider.GetRequiredService<IOptions<ExchangeConnectionOptions>>();

    public string ImpersonateUpn => ConnectionOptions.Value.ImpersonateUpn;


    public ExchangeService CreateTestService()
    {
        var options = ConnectionOptions.Value;

        return new ExchangeService
        {
            Credentials = new WebCredentials(options.UserName, options.Password),
            UseDefaultCredentials = false,
            PreAuthenticate = false,
            AcceptGzipEncoding = true,
            Url = new Uri(options.ServiceUrl),
            ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, options.ImpersonateUpn),
            ServerCertificateValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator,
        };
    }


    private static IServiceProvider BuildServiceProvider()
    {
        var configuration = new ConfigurationBuilder()
            // 
            .AddUserSecrets(Assembly.GetExecutingAssembly(), true)
            .Build();


        var collection = new ServiceCollection();
        collection.AddSingleton<IConfiguration>(configuration);

        collection.AddOptions<ExchangeConnectionOptions>().Bind(configuration.GetSection("Exchange"));


        return collection.BuildServiceProvider(true);
    }
}
