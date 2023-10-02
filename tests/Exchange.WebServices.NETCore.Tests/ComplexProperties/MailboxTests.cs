using Microsoft.Exchange.WebServices.Data;

namespace Exchange.WebServices.NETCore.Tests.ComplexProperties;

public class MailboxTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;


    public MailboxTests(ExchangeProvider provider)
    {
        _provider = provider;
    }


    [Fact]
    public void EqualityTest()
    {
        var a = new Mailbox("hello@world.com");
        var b = new Mailbox("world@hello.com");
        var c = new Mailbox("hello@world.com");

        // ReSharper disable EqualExpressionComparison
        Assert.True(a == a);
        Assert.False(a != a);
        Assert.False(a == null);
        Assert.False(null == a);
        Assert.False(a == b);
        Assert.True(a != b);

        Assert.False(a.Equals(b));
        Assert.False(a.Equals(null));
        Assert.True(Equals(a, a));
        Assert.False(Equals(a, b));
        // ReSharper restore EqualExpressionComparison

        Assert.True(a == c);
    }
}
