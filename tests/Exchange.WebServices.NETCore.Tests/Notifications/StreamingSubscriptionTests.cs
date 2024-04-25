using System.Diagnostics;

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
            [EventType.Created, EventType.Modified,]
        );

        using var connection = new StreamingSubscriptionConnection(service, [subscription,], 30);

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

    [Fact]
    public async Task StreamingEventsTest()
    {
        var service = _provider.CreateTestService();

        EventType[] events =
        [
            EventType.NewMail, EventType.Deleted, EventType.Modified, EventType.Moved, EventType.Copied,
            EventType.Created, EventType.FreeBusyChanged,
        ];

        var sub = await service.SubscribeToStreamingNotificationsOnAllFolders(eventTypes: events);

        using (var conn = new StreamingSubscriptionConnection(service, 1))
        {
            // Lifetime = one min
            conn.AddSubscription(sub);

            conn.OnNotificationEvent += (sender, args) => { Debugger.Break(); };

            // Never called. Should be called one min after calling Open();
            conn.OnDisconnect += (sender, args) => { Debugger.Break(); };
            conn.OnSubscriptionError += (sender, args) => { Debugger.Break(); };

            conn.Open();
        }
    }
}
