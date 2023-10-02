using JetBrains.Annotations;

namespace Exchange.WebServices.NETCore.Enumerations;

/// <summary>
/// Well known folder class types. Extracted from: https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxosfld/68a85898-84fe-43c4-b166-4711c13cdd61
/// </summary>
[PublicAPI]
public class WellKnownFolderClass
{
    /// <summary>
    /// Email messages or folders.
    /// </summary>
    public const string Note = "IPF.Note";

    /// <summary>
    /// Appointments and meetings.
    /// </summary>
    public const string Calendar = "IPF.Appointment";

    /// <summary>
    /// Contacts and distribution lists.
    /// </summary>
    public const string Contacts = "IPF.Contact";

    /// <summary>
    /// Contains Contact objects for the user's favorite contacts and instant messaging contacts.
    /// </summary>
    public const string QuickContacts = "IPF.Contact.MOC.QuickContacts";

    /// <summary>
    /// Contains Personal Distribution List objects of favorite contacts and instant messaging contacts.
    /// </summary>
    public const string ImContactsList = "IPF.Contact.MOC.ImContactList";

    /// <summary>
    /// Contains documents to be uploaded to a shared location.
    /// </summary>
    public const string DocumentLibraries = "IPF.ShortcutFolder";

    /// <summary>
    /// Contains Note objects.
    /// </summary>
    public const string StickyNote = "IPF.StickyNote";

    /// <summary>
    /// Contains Journal objects.
    /// </summary>
    public const string Journal = "IPF.Journal";

    /// <summary>
    /// Contains work items to complete.
    /// </summary>
    public const string Task = "IPF.Task";

    /// <summary>
    /// Contains Really Simple Syndication (RSS) feed messages.
    /// </summary>
    public const string RssFeed = "IPF.Note.OutlookHomepage";

    /// <summary>
    /// Contains folder associated information (FAI) messages that are used for persisting conversation actions.
    /// </summary>
    public const string ConversationActionSettings = "IPF.Configuration";
}
