using JetBrains.Annotations;

namespace Exchange.WebServices.NETCore.Enumerations;

/// <summary>
/// Well defined message classes. Extracted from https://learn.microsoft.com/en-us/office/vba/outlook/concepts/forms/item-types-and-message-classes
/// </summary>
[PublicAPI]
public class WellKnownMessageClass
{
    /// <summary>
    /// Journal entries
    /// </summary>
    public const string Activity = "IPM.Activity";

    /// <summary>
    /// Appointments
    /// </summary>
    public const string Appointment = "IPM.Appointment";

    /// <summary>
    /// Contacts
    /// </summary>
    public const string Contact = "IPM.Contact";

    /// <summary>
    /// Distribution lists
    /// </summary>
    public const string DistributionList = "IPM.DistList";

    /// <summary>
    /// Exception item of a recurrence series
    /// </summary>
    public const string OleClass = "IPM.OLE.Class";

    /// <summary>
    /// Documents
    /// </summary>
    public const string Document = "IPM.Document";

    /// <summary>
    /// Items for which the specified form cannot be found
    /// </summary>
    public const string Item = "IPM";

    /// <summary>
    /// Email messages
    /// </summary>
    public const string Note = "IPM.Note";

    /// <summary>
    /// Reports from the Internet Mail Connect (the Exchange Server gateway to the Internet)
    /// </summary>
    public const string ImcNotification = "IPM.Note.IMC.Notification";

    /// <summary>
    /// Out-of-office templates
    /// </summary>
    public const string OofTemplate = "IPM.Note.Rules.OofTemplate.Microsoft";

    /// <summary>
    /// Encrypted notes to other people
    /// </summary>
    public const string NoteSecure = "IPM.Note.Secure";

    /// <summary>
    /// Digitally signed notes to other people
    /// </summary>
    public const string NoteSecureSign = "IPM.Note.Secure.Sign";

    /// <summary>
    /// Posting notes in a folder
    /// </summary>
    public const string Post = "IPM.Post";

    /// <summary>
    /// Creating notes
    /// </summary>
    public const string StickyNote = "IPM.StickyNote";

    /// <summary>
    /// Message recall reports
    /// </summary>
    public const string RecallReport = "IPM.Recall.Report";

    /// <summary>
    /// Recalling sent messages from recipient Inboxes
    /// </summary>
    public const string OutlookRecall = "IPM.Outlook.Recall";

    /// <summary>
    /// Remote Mail message headers
    /// </summary>
    public const string Remote = "IPM.Remote";

    /// <summary>
    /// Editing rule reply templates
    /// </summary>
    public const string ReplyTemplate = "IPM.Note.Rules.ReplyTemplate.Microsoft";

    /// <summary>
    /// Reporting item status
    /// </summary>
    public const string Report = "IPM.Report";

    /// <summary>
    /// Resending a failed message
    /// </summary>
    public const string Resend = "IPM.Resend";

    /// <summary>
    /// Meeting cancellations
    /// </summary>
    public const string MeetingCanceled = "IPM.Schedule.Meeting.Canceled";

    /// <summary>
    /// Meeting requests
    /// </summary>
    public const string MeetingRequest = "IPM.Schedule.Meeting.Request";

    /// <summary>
    /// Responses to decline meeting requests
    /// </summary>
    public const string MeetingRespNeg = "IPM.Schedule.Meeting.Resp.Neg";

    /// <summary>
    /// Responses to accept meeting requests
    /// </summary>
    public const string MeetingRespPos = "IPM.Schedule.Meeting.Resp.Pos";

    /// <summary>
    /// Responses to tentatively accept meeting requests
    /// </summary>
    public const string MeetingRespTent = "IPM.Schedule.Meeting.Resp.Tent";

    /// <summary>
    /// Tasks
    /// </summary>
    public const string Task = "IPM.Task";

    /// <summary>
    /// Task requests
    /// </summary>
    public const string TaskRequest = "IPM.TaskRequest";

    /// <summary>
    /// Responses to accept task requests
    /// </summary>
    public const string TaskRequestAccept = "IPM.TaskRequest.Accept";

    /// <summary>
    /// Responses to decline task requests
    /// </summary>
    public const string TaskRequestDecline = "IPM.TaskRequest.Decline";

    /// <summary>
    /// Updates to requested tasks
    /// </summary>
    public const string TaskRequestUpdate = "IPM.TaskRequest.Update";
}
