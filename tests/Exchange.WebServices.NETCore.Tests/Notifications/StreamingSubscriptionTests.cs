using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        var folderIds = folders.Folders.ToList().Select(x => x.Id).ToList();

        var subscription = await service.SubscribeToStreamingNotifications(
            folderIds,
            new[]
            {
                EventType.Created, EventType.Modified,
            }
        );

        using var connection = new StreamingSubscriptionConnection(
            service,
            new[]
            {
                subscription,
            },
            30
        );

        connection.OnNotificationEvent += (sender, args) =>
        {
            //
            Debugger.Break();
        };

        connection.OnSubscriptionError += (sender, args) =>
        {
            //
            Debugger.Break();
        };

        connection.Open();
        Assert.True(connection.IsOpen);

        await Task.Delay(10_000);

        connection.Close();
    }
}
