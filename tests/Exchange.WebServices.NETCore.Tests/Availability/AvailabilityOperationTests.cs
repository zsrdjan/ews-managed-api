using System.Diagnostics;

using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Availability;

public class AvailabilityOperationTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;


    public AvailabilityOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }


    [Fact]
    public async Task GetRoomListTest()
    {
        var service = _provider.CreateTestService();

        var rooms = await service.GetRoomLists();

        Assert.Empty(rooms);
    }
}
