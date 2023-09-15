using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
/// Represents the results of a GetMailTips operation.
/// </summary>
[PublicAPI]
public sealed class GetMailTipsResults : ServiceResponse
{
    /// <summary>
    /// Initializes a new instance of the <see cref="GetMailTipsResults"/> class.
    /// </summary>
    internal GetMailTipsResults()
    {
    }

    /// <summary>
    /// Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        MailTipsResponses = new ServiceResponseCollection<MailTipsResponseMessage>();
        base.ReadElementsFromXml(reader);
        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

        if (!reader.IsEmptyElement)
        {
            // Because we don't have count of returned objects
            // test the element to determine if it is return object or EndElement
            reader.Read();
            while (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.MailTipsResponseMessageType))
            {
                var response = new MailTipsResponseMessage();
                response.LoadFromXml(reader, XmlElementNames.MailTipsResponseMessageType);
                MailTipsResponses.Add(response);
                reader.Read();
            }

            reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages);
        }
    }

    /// <summary>
    /// Gets a collection of MailTips responses for the requested recipients
    /// </summary>
    public ServiceResponseCollection<MailTipsResponseMessage> MailTipsResponses { get; internal set; }
}
