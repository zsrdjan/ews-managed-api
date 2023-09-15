using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Folders;

public class FolderOperationTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public FolderOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task FindFoldersTest()
    {
        var service = _provider.CreateTestService();

        var folders = await service.FindFolders(
            new FolderId(WellKnownFolderName.MsgFolderRoot),
            new FolderView(100, 0)
        );

        Assert.NotEmpty(folders);
    }

    //[Fact]
    //public async Task FindFoldersSearchFilterTest()
    //{
    //    var service = _provider.CreateTestService();

    //    var filter = new SearchFilter.SearchFilterCollection(
    //        LogicalOperator.Or,
    //        new SearchFilter.IsNotEqualTo(FolderSchema.DisplayName, "HelloWorld"),
    //        new SearchFilter.IsNotEqualTo(FolderSchema.DisplayName, "Test1234")
    //    );


    //    var folders = await service.FindFolders(
    //        new FolderId(WellKnownFolderName.ArchiveInbox),
    //        filter,
    //        new FolderView(200, 0)
    //    );

    //    Assert.NotEmpty(folders);
    //}
}
