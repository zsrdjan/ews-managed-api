using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.ComplexProperties;

public class FolderIdTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;


    public FolderIdTests(ExchangeProvider provider)
    {
        _provider = provider;
    }


    [Fact]
    public void EqualityTests()
    {
        var a = new FolderId(WellKnownFolderName.AdminAuditLogs, new Mailbox());
        var b = new FolderId(WellKnownFolderName.ArchiveInbox, new Mailbox());

        Assert.False(a.Equals(b));
        Assert.False(a == b);
    }

    [Fact]
    public void BrokenEquality()
    {
        var a = new FolderId(WellKnownFolderName.AdminAuditLogs);
        var b = new FolderId(WellKnownFolderName.ArchiveInbox);

        Assert.False(a.Equals(b));
    }

    [Fact]
    public async Task EqualityTest()
    {
        var service = _provider.CreateTestService();

        var folders = (await service.FindFolders(
            new FolderId(WellKnownFolderName.MsgFolderRoot),
            new FolderView(100, 0)
        )).ToList();

        Assert.True(folders.Count > 2);

        var a = folders[0].Id;
        var b = folders[1].Id;

        Assert.NotNull(a);
        Assert.NotNull(b);


        var c = new FolderId(a.UniqueId);

        // ReSharper disable EqualExpressionComparison
        Assert.True(a == a);
        Assert.True(a.Equals(a));
        Assert.False(a == null);
        Assert.False(null == a);

        Assert.False(a.Equals(b));
        Assert.False(a == b);
        // ReSharper restore EqualExpressionComparison

        Assert.True(a == c);
    }

    [Fact]
    public void MailboxFolderIdEqualityTest()
    {
        var a = new FolderId(WellKnownFolderName.ArchiveInbox, new Mailbox("hello@world.com"));
        var b = new FolderId(WellKnownFolderName.AdminAuditLogs, new Mailbox("world@hello.com"));
        var c = new FolderId(WellKnownFolderName.ArchiveInbox, new Mailbox("hello@world.com"));

        Assert.True(a.Equals(a));
        Assert.False(a == b);

        Assert.True(a == c);
    }
}
