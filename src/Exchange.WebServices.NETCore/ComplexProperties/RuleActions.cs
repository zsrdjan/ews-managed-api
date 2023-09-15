/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

using System.Collections.ObjectModel;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the set of actions available for a rule.
/// </summary>
[PublicAPI]
public sealed class RuleActions : ComplexProperty
{
    /// <summary>
    ///     SMS recipient address type.
    /// </summary>
    private const string MobileType = "MOBILE";

    /// <summary>
    ///     The CopyToFolder action.
    /// </summary>
    private FolderId _copyToFolder;

    /// <summary>
    ///     The Delete action.
    /// </summary>
    private bool _delete;

    /// <summary>
    ///     The MarkImportance action.
    /// </summary>
    private Importance? _markImportance;

    /// <summary>
    ///     The MarkAsRead action.
    /// </summary>
    private bool _markAsRead;

    /// <summary>
    ///     The MoveToFolder action.
    /// </summary>
    private FolderId _moveToFolder;

    /// <summary>
    ///     The PermanentDelete action.
    /// </summary>
    private bool _permanentDelete;

    /// <summary>
    ///     The ServerReplyWithMessage action.
    /// </summary>
    private ItemId _serverReplyWithMessage;

    /// <summary>
    ///     The StopProcessingRules action.
    /// </summary>
    private bool _stopProcessingRules;

    /// <summary>
    ///     Initializes a new instance of the <see cref="RulePredicates" /> class.
    /// </summary>
    internal RuleActions()
    {
        AssignCategories = new StringList();
        ForwardAsAttachmentToRecipients = new EmailAddressCollection(XmlElementNames.Address);
        ForwardToRecipients = new EmailAddressCollection(XmlElementNames.Address);
        RedirectToRecipients = new EmailAddressCollection(XmlElementNames.Address);
        SendSMSAlertToRecipients = new Collection<MobilePhone>();
    }

    /// <summary>
    ///     Gets the categories that should be stamped on incoming messages.
    ///     To disable stamping incoming messages with categories, set
    ///     AssignCategories to null.
    /// </summary>
    public StringList AssignCategories { get; }

    /// <summary>
    ///     Gets or sets the Id of the folder incoming messages should be copied to.
    ///     To disable copying incoming messages to a folder, set CopyToFolder to null.
    /// </summary>
    public FolderId? CopyToFolder
    {
        get => _copyToFolder;
        set => SetFieldValue(ref _copyToFolder, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages should be
    ///     automatically moved to the Deleted Items folder.
    /// </summary>
    public bool Delete
    {
        get => _delete;
        set => SetFieldValue(ref _delete, value);
    }

    /// <summary>
    ///     Gets the e-mail addresses to which incoming messages should be
    ///     forwarded as attachments. To disable forwarding incoming messages
    ///     as attachments, empty the ForwardAsAttachmentToRecipients list.
    /// </summary>
    public EmailAddressCollection ForwardAsAttachmentToRecipients { get; }

    /// <summary>
    ///     Gets the e-mail addresses to which incoming messages should be forwarded.
    ///     To disable forwarding incoming messages, empty the ForwardToRecipients list.
    /// </summary>
    public EmailAddressCollection ForwardToRecipients { get; }

    /// <summary>
    ///     Gets or sets the importance that should be stamped on incoming
    ///     messages. To disable the stamping of incoming messages with an
    ///     importance, set MarkImportance to null.
    /// </summary>
    public Importance? MarkImportance
    {
        get => _markImportance;
        set => SetFieldValue(ref _markImportance, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages should be
    ///     marked as read.
    /// </summary>
    public bool MarkAsRead
    {
        get => _markAsRead;
        set => SetFieldValue(ref _markAsRead, value);
    }

    /// <summary>
    ///     Gets or sets the Id of the folder to which incoming messages should be
    ///     moved. To disable the moving of incoming messages to a folder, set
    ///     CopyToFolder to null.
    /// </summary>
    public FolderId? MoveToFolder
    {
        get => _moveToFolder;
        set => SetFieldValue(ref _moveToFolder, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages should be
    ///     permanently deleted. When a message is permanently deleted, it is never
    ///     saved into the recipient's mailbox. To delete a message after it has
    ///     been saved into the recipient's mailbox, use the Delete action.
    /// </summary>
    public bool PermanentDelete
    {
        get => _permanentDelete;
        set => SetFieldValue(ref _permanentDelete, value);
    }

    /// <summary>
    ///     Gets the e-mail addresses to which incoming messages should be
    ///     redirecteded. To disable redirection of incoming messages, empty
    ///     the RedirectToRecipients list. Unlike forwarded mail, redirected mail
    ///     maintains the original sender and recipients.
    /// </summary>
    public EmailAddressCollection RedirectToRecipients { get; }

    /// <summary>
    ///     Gets the phone numbers to which an SMS alert should be sent. To disable
    ///     sending SMS alerts for incoming messages, empty the
    ///     SendSMSAlertToRecipients list.
    /// </summary>
    public Collection<MobilePhone> SendSMSAlertToRecipients { get; private set; }

    /// <summary>
    ///     Gets or sets the Id of the template message that should be sent
    ///     as a reply to incoming messages. To disable automatic replies, set
    ///     ServerReplyWithMessage to null.
    /// </summary>
    public ItemId? ServerReplyWithMessage
    {
        get => _serverReplyWithMessage;
        set => SetFieldValue(ref _serverReplyWithMessage, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether subsequent rules should be
    ///     evaluated.
    /// </summary>
    public bool StopProcessingRules
    {
        get => _stopProcessingRules;
        set => SetFieldValue(ref _stopProcessingRules, value);
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.AssignCategories:
            {
                AssignCategories.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.CopyToFolder:
            {
                reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.FolderId);
                _copyToFolder = new FolderId();
                _copyToFolder.LoadFromXml(reader, XmlElementNames.FolderId);
                reader.ReadEndElement(XmlNamespace.NotSpecified, XmlElementNames.CopyToFolder);
                return true;
            }
            case XmlElementNames.Delete:
            {
                _delete = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.ForwardAsAttachmentToRecipients:
            {
                ForwardAsAttachmentToRecipients.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ForwardToRecipients:
            {
                ForwardToRecipients.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.MarkImportance:
            {
                _markImportance = reader.ReadElementValue<Importance>();
                return true;
            }
            case XmlElementNames.MarkAsRead:
            {
                _markAsRead = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.MoveToFolder:
            {
                reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.FolderId);
                _moveToFolder = new FolderId();
                _moveToFolder.LoadFromXml(reader, XmlElementNames.FolderId);
                reader.ReadEndElement(XmlNamespace.NotSpecified, XmlElementNames.MoveToFolder);
                return true;
            }
            case XmlElementNames.PermanentDelete:
            {
                _permanentDelete = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.RedirectToRecipients:
            {
                RedirectToRecipients.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.SendSMSAlertToRecipients:
            {
                var smsRecipientCollection = new EmailAddressCollection(XmlElementNames.Address);
                smsRecipientCollection.LoadFromXml(reader, reader.LocalName);
                SendSMSAlertToRecipients =
                    ConvertSmsRecipientsFromEmailAddressCollectionToMobilePhoneCollection(smsRecipientCollection);
                return true;
            }
            case XmlElementNames.ServerReplyWithMessage:
            {
                _serverReplyWithMessage = new ItemId();
                _serverReplyWithMessage.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.StopProcessingRules:
            {
                _stopProcessingRules = reader.ReadElementValue<bool>();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (AssignCategories.Count > 0)
        {
            AssignCategories.WriteToXml(writer, XmlElementNames.AssignCategories);
        }

        if (CopyToFolder != null)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.CopyToFolder);
            CopyToFolder.WriteToXml(writer);
            writer.WriteEndElement();
        }

        if (Delete)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Delete, Delete);
        }

        if (ForwardAsAttachmentToRecipients.Count > 0)
        {
            ForwardAsAttachmentToRecipients.WriteToXml(writer, XmlElementNames.ForwardAsAttachmentToRecipients);
        }

        if (ForwardToRecipients.Count > 0)
        {
            ForwardToRecipients.WriteToXml(writer, XmlElementNames.ForwardToRecipients);
        }

        if (MarkImportance.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MarkImportance, MarkImportance.Value);
        }

        if (MarkAsRead)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MarkAsRead, MarkAsRead);
        }

        if (MoveToFolder != null)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.MoveToFolder);
            MoveToFolder.WriteToXml(writer);
            writer.WriteEndElement();
        }

        if (PermanentDelete)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PermanentDelete, PermanentDelete);
        }

        if (RedirectToRecipients.Count > 0)
        {
            RedirectToRecipients.WriteToXml(writer, XmlElementNames.RedirectToRecipients);
        }

        if (SendSMSAlertToRecipients.Count > 0)
        {
            var emailCollection =
                ConvertSmsRecipientsFromMobilePhoneCollectionToEmailAddressCollection(SendSMSAlertToRecipients);
            emailCollection.WriteToXml(writer, XmlElementNames.SendSMSAlertToRecipients);
        }

        if (ServerReplyWithMessage != null)
        {
            ServerReplyWithMessage.WriteToXml(writer, XmlElementNames.ServerReplyWithMessage);
        }

        if (StopProcessingRules)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.StopProcessingRules, StopProcessingRules);
        }
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void InternalValidate()
    {
        base.InternalValidate();
        EwsUtilities.ValidateParam(ForwardAsAttachmentToRecipients);
        EwsUtilities.ValidateParam(ForwardToRecipients);
        EwsUtilities.ValidateParam(RedirectToRecipients);
        foreach (var sendSmsAlertToRecipient in SendSMSAlertToRecipients)
        {
            EwsUtilities.ValidateParam(sendSmsAlertToRecipient, "SendSMSAlertToRecipients");
        }
    }

    /// <summary>
    ///     Convert the SMS recipient list from EmailAddressCollection type to MobilePhone collection type.
    /// </summary>
    /// <param name="emailCollection">Recipient list in EmailAddressCollection type.</param>
    /// <returns>A MobilePhone collection object containing all SMS recipient in MobilePhone type. </returns>
    private static Collection<MobilePhone> ConvertSmsRecipientsFromEmailAddressCollectionToMobilePhoneCollection(
        EmailAddressCollection emailCollection
    )
    {
        var mobilePhoneCollection = new Collection<MobilePhone>();
        foreach (var emailAddress in emailCollection)
        {
            mobilePhoneCollection.Add(new MobilePhone(emailAddress.Name, emailAddress.Address));
        }

        return mobilePhoneCollection;
    }

    /// <summary>
    ///     Convert the SMS recipient list from MobilePhone collection type to EmailAddressCollection type.
    /// </summary>
    /// <param name="recipientCollection">Recipient list in a MobilePhone collection type.</param>
    /// <returns>An EmailAddressCollection object containing recipients with "MOBILE" address type. </returns>
    private static EmailAddressCollection ConvertSmsRecipientsFromMobilePhoneCollectionToEmailAddressCollection(
        Collection<MobilePhone> recipientCollection
    )
    {
        var emailCollection = new EmailAddressCollection(XmlElementNames.Address);
        foreach (var recipient in recipientCollection)
        {
            var emailAddress = new EmailAddress(recipient.Name, recipient.PhoneNumber, MobileType);
            emailCollection.Add(emailAddress);
        }

        return emailCollection;
    }
}
