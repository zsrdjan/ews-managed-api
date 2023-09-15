using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
/// Defines the types of requested mail tips.
/// </summary>
[PublicAPI]
public enum MailTipsRequested
{
    /// <summary>
    /// Represents all available mail tips.
    /// </summary>
    All,

    /// <summary>
    /// Represents the Out of Office (OOF) message.
    /// </summary>
    OutOfOfficeMessage,

    /// <summary>
    /// Represents the status for a mailbox that is full.
    /// </summary>
    MailboxFullStatus,

    /// <summary>
    /// Represents a custom mail tip.
    /// </summary>
    CustomMailTip,

    /// <summary>
    /// Represents the count of external members.
    /// </summary>
    ExternalMemberCount,

    /// <summary>
    /// Represents the count of all members.
    /// </summary>
    TotalMemberCount,

    /// <summary>
    /// Represents the maximum message size a recipient can accept.
    /// </summary>
    MaxMessageSize,

    /// <summary>
    /// Indicates whether delivery restrictions will prevent the sender's message from reaching the recipient.
    /// </summary>
    DeliveryRestriction,

    /// <summary>
    /// Indicates whether the sender's message will be reviewed by a moderator.
    /// </summary>
    ModerationStatus,

    /// <summary>
    /// Indicates whether the recipient is invalid.
    /// </summary>
    InvalidRecipient,
}

/// <summary>
/// Represents a GetMailTips request.
/// </summary>
internal sealed class GetMailTipsRequest : SimpleServiceRequestBase
{
    //https://msdn.microsoft.com/en-us/library/office/dd877060(v=exchg.140).aspx [GetMailTips Operation][2010]
    //https://msdn.microsoft.com/en-us/library/office/dd877060(v=exchg.150).aspx [GetMailTips Operation][2013]

    /// <summary>
    /// Initializes a new instance of the <see cref="GetMailTipsRequest"/> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal GetMailTipsRequest(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    /// Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.GetMailTips;
    }

    /// <summary>Writes XML elements.</summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        SendingAs.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.SendingAs);

        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Recipients);
        foreach (var mailbox in Recipients)
        {
            mailbox.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.Mailbox);
        }

        writer.WriteEndElement(); // </Recipients>

        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MailTipsRequested, MailTipsRequested);
    }

    /// <summary>Gets the name of the response XML element.</summary>
    /// <returns>XML element name</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.GetMailTipsResponse;
    }

    /// <summary>Parses the response.</summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Response object.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader)
    {
        var serviceResponse = new GetMailTipsResults();
        serviceResponse.LoadFromXml(reader, XmlElementNames.GetMailTipsResponse);
        return serviceResponse;
    }

    /// <summary>Gets the request version.</summary>
    /// <returns>Earliest Exchange version in which this request is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2010;
    }

    /// <summary>Executes this request.</summary>
    /// <returns>Service response.</returns>
    internal async Task<GetMailTipsResults> Execute(CancellationToken token)
    {
        return await InternalExecuteAsync<GetMailTipsResults>(token).ConfigureAwait(false);
    }

    /// <summary>Gets or sets the attendees.</summary>
    public EmailAddress SendingAs { get; set; }

    /// <summary>Gets or sets the requested MailTips.</summary>
    public MailTipsRequested MailTipsRequested { get; set; }

    /// <summary>
    /// Gets or sets who are the recipients/targets whose MailTips we are interested in.
    /// </summary>
    public Mailbox[] Recipients { get; set; }
}
