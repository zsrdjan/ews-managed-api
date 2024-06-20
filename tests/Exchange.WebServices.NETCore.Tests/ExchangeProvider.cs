// #define TRACING

using System.Reflection;

using Exchange.WebServices.NETCore.Tests.Utility;

using JetBrains.Annotations;

using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Web.TokenCacheProviders;
using Microsoft.Identity.Web.TokenCacheProviders.InMemory;

namespace Exchange.WebServices.NETCore.Tests;

[PublicAPI]
public class ExchangeConnectionOptions
{
    public required string ServiceUrl { get; set; }

    public required string UserName { get; set; }

    public required string Password { get; set; }

    public required string ImpersonationUpn { get; set; }
}

[PublicAPI]
public class OutlookConnectionOptions
{
    public required string Url { get; set; }

    public required string TenantId { get; set; }

    public required string ClientId { get; set; }

    public required string ClientSecret { get; set; }

    public required string ImpersonationUpn { get; set; }
}

[PublicAPI]
public class ExchangeProvider
{
    protected readonly IServiceProvider _provider = BuildServiceProvider();


    public IOptions<ExchangeConnectionOptions> ConnectionOptions =>
        _provider.GetRequiredService<IOptions<ExchangeConnectionOptions>>();

    public OutlookConnectionOptions OutlookConnectionOptions =>
        _provider.GetRequiredService<IOptions<OutlookConnectionOptions>>().Value;

    public string ImpersonateUpn => ConnectionOptions.Value.ImpersonationUpn;


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
            ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, options.ImpersonationUpn),
            ServerCertificateValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator,
            SendClientLatencies = true,
#if TRACING
            TraceEnabled = true,
            TraceFlags = TraceFlags.All,
            TraceListener = new EwsTraceListener(),
#endif
        };
    }

    public T GetRequiredService<T>()
        where T : notnull
    {
        return _provider.GetRequiredService<T>();
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
        collection.AddOptions<OutlookConnectionOptions>().Bind(configuration.GetSection("Outlook"));

        // Add memory cache for MSALTokenCacheProvider
        collection.AddSingleton<IMemoryCache, MemoryCache>();
        collection.AddSingleton<IMsalTokenCacheProvider, MsalMemoryTokenCacheProvider>();


        return collection.BuildServiceProvider(true);
    }
}
