namespace Exchange.WebServices.NETCore.Tests.Folders;

public class FolderOperationTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public FolderOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task Test()
    {
        var service = _provider.CreateTestService();


        //
    }
}
