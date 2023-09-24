using Microsoft.Exchange.WebServices.Data;

namespace Exchange.WebServices.NETCore.Tests.Core.Schema;

public class SchemaTests
{
    [Fact]
    public void InstantiationTest()
    {
        // Crude check to test for schema instantiation failures.
        // This can occur if the static field locations are reordered in the file.

#pragma warning disable IDE0059 // Unnecessary assignment of a value
        var a = AppointmentSchema.StartTimeZone;
        var b = ContactGroupSchema.DisplayName;
        var c = ContactSchema.DisplayName;
        var d = ConversationSchema.Categories;
        var e = EmailMessageSchema.ToRecipients;
        var f = FolderSchema.DisplayName;
        var g = ItemSchema.DisplayCc;
        var h = MeetingCancellationSchema.Start;
        var i = MeetingMessageSchema.AssociatedAppointmentId;
        var j = MeetingRequestSchema.MeetingRequestType;
        var k = MeetingResponseSchema.Start;
        var l = PersonaSchema.DisplayName;
        var m = PostItemSchema.PostedTime;
        var o = SearchFolderSchema.SearchParameters;
        var p = TaskSchema.AssignedTime;
#pragma warning restore IDE0059
    }
}
