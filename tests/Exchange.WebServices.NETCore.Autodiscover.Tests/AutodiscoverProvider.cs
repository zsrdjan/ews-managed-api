using System.Reflection;

using JetBrains.Annotations;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace Exchange.WebServices.NETCore.Autodiscover.Tests;

[PublicAPI]
public class AutodiscoverConnectionOptions
{
    public required string UserName { get; set; }

    public required string Password { get; set; }
}

[PublicAPI]
public class AutodiscoverProvider
{
    protected readonly IServiceProvider _provider = BuildServiceProvider();

    public IOptions<AutodiscoverConnectionOptions> ConnectionOptions =>
        _provider.GetRequiredService<IOptions<AutodiscoverConnectionOptions>>();


    private static IServiceProvider BuildServiceProvider()
    {
        var configuration = new ConfigurationBuilder()
            // 
            .AddUserSecrets(Assembly.GetExecutingAssembly(), true)
            .Build();


        var collection = new ServiceCollection();
        collection.AddSingleton<IConfiguration>(configuration);

        collection.AddOptions<AutodiscoverConnectionOptions>().Bind(configuration.GetSection("Autodiscover"));

        return collection.BuildServiceProvider(true);
    }
}
