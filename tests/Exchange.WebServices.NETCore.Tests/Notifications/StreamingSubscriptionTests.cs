using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Notifications;

public class StreamingSubscriptionTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public StreamingSubscriptionTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task StreamingTest()
    {
        var service = _provider.CreateTestService();

        var folders = await service.FindFolders(
            new FolderId(WellKnownFolderName.MsgFolderRoot),
            new FolderView(100, 0)
        );

        var folderIds = folders.Folders.ToList().Select(x => x.Id);

        var subscription = await service.SubscribeToStreamingNotifications(folderIds);

        var connection = new StreamingSubscriptionConnection(
            service,
            new[]
            {
                subscription,
            },
            30
        );

        connection.Open();
    }
}
