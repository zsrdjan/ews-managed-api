using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Items;

public class ItemOperationTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public ItemOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task ItemSearchFilterTest()
    {
        var service = _provider.CreateTestService();

        _ = await Folder.Bind(service, WellKnownFolderName.Inbox);

        // The search filter to get unread email.
        var filter = new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
        );
        var view = new ItemView(1);

        var items = await service.FindItems(WellKnownFolderName.Inbox, filter, view);
        Assert.NotEmpty(items);
    }
}
