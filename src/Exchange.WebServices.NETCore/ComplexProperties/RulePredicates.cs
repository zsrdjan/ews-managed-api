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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the set of conditions and exceptions available for a rule.
/// </summary>
[PublicAPI]
public sealed class RulePredicates : ComplexProperty
{
    /// <summary>
    ///     The FlaggedForAction predicate.
    /// </summary>
    private FlaggedForAction? _flaggedForAction;

    /// <summary>
    ///     The HasAttachments predicate.
    /// </summary>
    private bool _hasAttachments;

    /// <summary>
    ///     The Importance predicate.
    /// </summary>
    private Importance? _importance;

    /// <summary>
    ///     The IsApprovalRequest predicate.
    /// </summary>
    private bool _isApprovalRequest;

    /// <summary>
    ///     The IsAutomaticForward predicate.
    /// </summary>
    private bool _isAutomaticForward;

    /// <summary>
    ///     The IsAutomaticReply predicate.
    /// </summary>
    private bool _isAutomaticReply;

    /// <summary>
    ///     The IsEncrypted predicate.
    /// </summary>
    private bool _isEncrypted;

    /// <summary>
    ///     The IsMeetingRequest predicate.
    /// </summary>
    private bool _isMeetingRequest;

    /// <summary>
    ///     The IsMeetingResponse predicate.
    /// </summary>
    private bool _isMeetingResponse;

    /// <summary>
    ///     The IsNDR predicate.
    /// </summary>
    private bool _isNonDeliveryReport;

    /// <summary>
    ///     The IsPermissionControlled predicate.
    /// </summary>
    private bool _isPermissionControlled;

    /// <summary>
    ///     The IsReadReceipt  predicate.
    /// </summary>
    private bool _isReadReceipt;

    /// <summary>
    ///     The IsSigned predicate.
    /// </summary>
    private bool _isSigned;

    /// <summary>
    ///     The IsVoicemail predicate.
    /// </summary>
    private bool _isVoicemail;

    /// <summary>
    ///     The NotSentToMe predicate.
    /// </summary>
    private bool _notSentToMe;

    /// <summary>
    ///     The Sensitivity predicate.
    /// </summary>
    private Sensitivity? _sensitivity;

    /// <summary>
    ///     SentCcMe predicate.
    /// </summary>
    private bool _sentCcMe;

    /// <summary>
    ///     The SentOnlyToMe predicate.
    /// </summary>
    private bool _sentOnlyToMe;

    /// <summary>
    ///     The SentToMe predicate.
    /// </summary>
    private bool _sentToMe;

    /// <summary>
    ///     The SentToOrCcMe predicate.
    /// </summary>
    private bool _sentToOrCcMe;

    /// <summary>
    ///     Gets the categories that an incoming message should be stamped with
    ///     for the condition or exception to apply. To disable this predicate,
    ///     empty the list.
    /// </summary>
    public StringList Categories { get; }

    /// <summary>
    ///     Gets the strings that should appear in the body of incoming messages
    ///     for the condition or exception to apply.
    ///     To disable this predicate, empty the list.
    /// </summary>
    public StringList ContainsBodyStrings { get; }

    /// <summary>
    ///     Gets the strings that should appear in the headers of incoming messages
    ///     for the condition or exception to apply. To disable this predicate, empty
    ///     the list.
    /// </summary>
    public StringList ContainsHeaderStrings { get; }

    /// <summary>
    ///     Gets the strings that should appear in either the To or Cc fields of
    ///     incoming messages for the condition or exception to apply. To disable this
    ///     predicate, empty the list.
    /// </summary>
    public StringList ContainsRecipientStrings { get; }

    /// <summary>
    ///     Gets the strings that should appear in the From field of incoming messages
    ///     for the condition or exception to apply. To disable this predicate, empty
    ///     the list.
    /// </summary>
    public StringList ContainsSenderStrings { get; }

    /// <summary>
    ///     Gets the strings that should appear in either the body or the subject
    ///     of incoming messages for the condition or exception to apply.
    ///     To disable this predicate, empty the list.
    /// </summary>
    public StringList ContainsSubjectOrBodyStrings { get; }

    /// <summary>
    ///     Gets the strings that should appear in the subject of incoming messages
    ///     for the condition or exception to apply. To disable this predicate,
    ///     empty the list.
    /// </summary>
    public StringList ContainsSubjectStrings { get; }

    /// <summary>
    ///     Gets or sets the flag for action value that should appear on incoming
    ///     messages for the condition or exception to apply. To disable this
    ///     predicate, set it to null.
    /// </summary>
    public FlaggedForAction? FlaggedForAction
    {
        get => _flaggedForAction;
        set => SetFieldValue(ref _flaggedForAction, value);
    }

    /// <summary>
    ///     Gets the e-mail addresses of the senders of incoming messages for the
    ///     condition or exception to apply. To disable this predicate, empty the
    ///     list.
    /// </summary>
    public EmailAddressCollection FromAddresses { get; }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must have
    ///     attachments for the condition or exception to apply.
    /// </summary>
    public bool HasAttachments
    {
        get => _hasAttachments;
        set => SetFieldValue(ref _hasAttachments, value);
    }

    /// <summary>
    ///     Gets or sets the importance that should be stamped on incoming messages
    ///     for the condition or exception to apply. To disable this predicate, set
    ///     it to null.
    /// </summary>
    public Importance? Importance
    {
        get => _importance;
        set => SetFieldValue(ref _importance, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     approval requests for the condition or exception to apply.
    /// </summary>
    public bool IsApprovalRequest
    {
        get => _isApprovalRequest;
        set => SetFieldValue(ref _isApprovalRequest, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     automatic forwards for the condition or exception to apply.
    /// </summary>
    public bool IsAutomaticForward
    {
        get => _isAutomaticForward;
        set => SetFieldValue(ref _isAutomaticForward, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     automatic replies for the condition or exception to apply.
    /// </summary>
    public bool IsAutomaticReply
    {
        get => _isAutomaticReply;
        set => SetFieldValue(ref _isAutomaticReply, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     S/MIME encrypted for the condition or exception to apply.
    /// </summary>
    public bool IsEncrypted
    {
        get => _isEncrypted;
        set => SetFieldValue(ref _isEncrypted, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     meeting requests for the condition or exception to apply.
    /// </summary>
    public bool IsMeetingRequest
    {
        get => _isMeetingRequest;
        set => SetFieldValue(ref _isMeetingRequest, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     meeting responses for the condition or exception to apply.
    /// </summary>
    public bool IsMeetingResponse
    {
        get => _isMeetingResponse;
        set => SetFieldValue(ref _isMeetingResponse, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     non-delivery reports (NDR) for the condition or exception to apply.
    /// </summary>
    public bool IsNonDeliveryReport
    {
        get => _isNonDeliveryReport;
        set => SetFieldValue(ref _isNonDeliveryReport, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     permission controlled (RMS protected) for the condition or exception
    ///     to apply.
    /// </summary>
    public bool IsPermissionControlled
    {
        get => _isPermissionControlled;
        set => SetFieldValue(ref _isPermissionControlled, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     S/MIME signed for the condition or exception to apply.
    /// </summary>
    public bool IsSigned
    {
        get => _isSigned;
        set => SetFieldValue(ref _isSigned, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     voice mails for the condition or exception to apply.
    /// </summary>
    public bool IsVoicemail
    {
        get => _isVoicemail;
        set => SetFieldValue(ref _isVoicemail, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether incoming messages must be
    ///     read receipts for the condition or exception to apply.
    /// </summary>
    public bool IsReadReceipt
    {
        get => _isReadReceipt;
        set => SetFieldValue(ref _isReadReceipt, value);
    }

    /// <summary>
    ///     Gets the e-mail account names from which incoming messages must have
    ///     been aggregated for the condition or exception to apply. To disable
    ///     this predicate, empty the list.
    /// </summary>
    public StringList FromConnectedAccounts { get; }

    /// <summary>
    ///     Gets the item classes that must be stamped on incoming messages for
    ///     the condition or exception to apply. To disable this predicate,
    ///     empty the list.
    /// </summary>
    public StringList ItemClasses { get; }

    /// <summary>
    ///     Gets the message classifications that must be stamped on incoming messages
    ///     for the condition or exception to apply. To disable this predicate,
    ///     empty the list.
    /// </summary>
    public StringList MessageClassifications { get; }

    /// <summary>
    ///     Gets or sets a value indicating whether the owner of the mailbox must
    ///     NOT be a To recipient of the incoming messages for the condition or
    ///     exception to apply.
    /// </summary>
    public bool NotSentToMe
    {
        get => _notSentToMe;
        set => SetFieldValue(ref _notSentToMe, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the owner of the mailbox must be
    ///     a Cc recipient of incoming messages for the condition or exception to apply.
    /// </summary>
    public bool SentCcMe
    {
        get => _sentCcMe;
        set => SetFieldValue(ref _sentCcMe, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the owner of the mailbox must be
    ///     the only To recipient of incoming messages for the condition or exception
    ///     to apply.
    /// </summary>
    public bool SentOnlyToMe
    {
        get => _sentOnlyToMe;
        set => SetFieldValue(ref _sentOnlyToMe, value);
    }

    /// <summary>
    ///     Gets the e-mail addresses incoming messages must have been sent to for
    ///     the condition or exception to apply. To disable this predicate, empty
    ///     the list.
    /// </summary>
    public EmailAddressCollection SentToAddresses { get; }

    /// <summary>
    ///     Gets or sets a value indicating whether the owner of the mailbox must be
    ///     a To recipient of incoming messages for the condition or exception to apply.
    /// </summary>
    public bool SentToMe
    {
        get => _sentToMe;
        set => SetFieldValue(ref _sentToMe, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the owner of the mailbox must be
    ///     either a To or Cc recipient of incoming messages for the condition or
    ///     exception to apply.
    /// </summary>
    public bool SentToOrCcMe
    {
        get => _sentToOrCcMe;
        set => SetFieldValue(ref _sentToOrCcMe, value);
    }

    /// <summary>
    ///     Gets or sets the sensitivity that must be stamped on incoming messages
    ///     for the condition or exception to apply. To disable this predicate, set it
    ///     to null.
    /// </summary>
    public Sensitivity? Sensitivity
    {
        get => _sensitivity;
        set => SetFieldValue(ref _sensitivity, value);
    }

    /// <summary>
    ///     Gets the date range within which incoming messages must have been received
    ///     for the condition or exception to apply. To disable this predicate, set both
    ///     its Start and End properties to null.
    /// </summary>
    public RulePredicateDateRange WithinDateRange { get; }

    /// <summary>
    ///     Gets the minimum and maximum sizes incoming messages must have for the
    ///     condition or exception to apply. To disable this predicate, set both its
    ///     MinimumSize and MaximumSize properties to null.
    /// </summary>
    public RulePredicateSizeRange WithinSizeRange { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="RulePredicates" /> class.
    /// </summary>
    internal RulePredicates()
    {
        Categories = new StringList();
        ContainsBodyStrings = new StringList();
        ContainsHeaderStrings = new StringList();
        ContainsRecipientStrings = new StringList();
        ContainsSenderStrings = new StringList();
        ContainsSubjectOrBodyStrings = new StringList();
        ContainsSubjectStrings = new StringList();
        FromAddresses = new EmailAddressCollection(XmlElementNames.Address);
        FromConnectedAccounts = new StringList();
        ItemClasses = new StringList();
        MessageClassifications = new StringList();
        SentToAddresses = new EmailAddressCollection(XmlElementNames.Address);
        WithinDateRange = new RulePredicateDateRange();
        WithinSizeRange = new RulePredicateSizeRange();
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
            case XmlElementNames.Categories:
            {
                Categories.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ContainsBodyStrings:
            {
                ContainsBodyStrings.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ContainsHeaderStrings:
            {
                ContainsHeaderStrings.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ContainsRecipientStrings:
            {
                ContainsRecipientStrings.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ContainsSenderStrings:
            {
                ContainsSenderStrings.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ContainsSubjectOrBodyStrings:
            {
                ContainsSubjectOrBodyStrings.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ContainsSubjectStrings:
            {
                ContainsSubjectStrings.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.FlaggedForAction:
            {
                _flaggedForAction = reader.ReadElementValue<FlaggedForAction>();
                return true;
            }
            case XmlElementNames.FromAddresses:
            {
                FromAddresses.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.FromConnectedAccounts:
            {
                FromConnectedAccounts.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.HasAttachments:
            {
                _hasAttachments = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.Importance:
            {
                _importance = reader.ReadElementValue<Importance>();
                return true;
            }
            case XmlElementNames.IsApprovalRequest:
            {
                _isApprovalRequest = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsAutomaticForward:
            {
                _isAutomaticForward = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsAutomaticReply:
            {
                _isAutomaticReply = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsEncrypted:
            {
                _isEncrypted = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsMeetingRequest:
            {
                _isMeetingRequest = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsMeetingResponse:
            {
                _isMeetingResponse = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsNDR:
            {
                _isNonDeliveryReport = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsPermissionControlled:
            {
                _isPermissionControlled = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsSigned:
            {
                _isSigned = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsVoicemail:
            {
                _isVoicemail = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsReadReceipt:
            {
                _isReadReceipt = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.ItemClasses:
            {
                ItemClasses.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.MessageClassifications:
            {
                MessageClassifications.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.NotSentToMe:
            {
                _notSentToMe = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.SentCcMe:
            {
                _sentCcMe = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.SentOnlyToMe:
            {
                _sentOnlyToMe = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.SentToAddresses:
            {
                SentToAddresses.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.SentToMe:
            {
                _sentToMe = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.SentToOrCcMe:
            {
                _sentToOrCcMe = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.Sensitivity:
            {
                _sensitivity = reader.ReadElementValue<Sensitivity>();
                return true;
            }
            case XmlElementNames.WithinDateRange:
            {
                WithinDateRange.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.WithinSizeRange:
            {
                WithinSizeRange.LoadFromXml(reader, reader.LocalName);
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
        if (Categories.Count > 0)
        {
            Categories.WriteToXml(writer, XmlElementNames.Categories);
        }

        if (ContainsBodyStrings.Count > 0)
        {
            ContainsBodyStrings.WriteToXml(writer, XmlElementNames.ContainsBodyStrings);
        }

        if (ContainsHeaderStrings.Count > 0)
        {
            ContainsHeaderStrings.WriteToXml(writer, XmlElementNames.ContainsHeaderStrings);
        }

        if (ContainsRecipientStrings.Count > 0)
        {
            ContainsRecipientStrings.WriteToXml(writer, XmlElementNames.ContainsRecipientStrings);
        }

        if (ContainsSenderStrings.Count > 0)
        {
            ContainsSenderStrings.WriteToXml(writer, XmlElementNames.ContainsSenderStrings);
        }

        if (ContainsSubjectOrBodyStrings.Count > 0)
        {
            ContainsSubjectOrBodyStrings.WriteToXml(writer, XmlElementNames.ContainsSubjectOrBodyStrings);
        }

        if (ContainsSubjectStrings.Count > 0)
        {
            ContainsSubjectStrings.WriteToXml(writer, XmlElementNames.ContainsSubjectStrings);
        }

        if (FlaggedForAction.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FlaggedForAction, FlaggedForAction.Value);
        }

        if (FromAddresses.Count > 0)
        {
            FromAddresses.WriteToXml(writer, XmlElementNames.FromAddresses);
        }

        if (FromConnectedAccounts.Count > 0)
        {
            FromConnectedAccounts.WriteToXml(writer, XmlElementNames.FromConnectedAccounts);
        }

        if (HasAttachments)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.HasAttachments, HasAttachments);
        }

        if (Importance.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Importance, Importance.Value);
        }

        if (IsApprovalRequest)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsApprovalRequest, IsApprovalRequest);
        }

        if (IsAutomaticForward)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsAutomaticForward, IsAutomaticForward);
        }

        if (IsAutomaticReply)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsAutomaticReply, IsAutomaticReply);
        }

        if (IsEncrypted)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsEncrypted, IsEncrypted);
        }

        if (IsMeetingRequest)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsMeetingRequest, IsMeetingRequest);
        }

        if (IsMeetingResponse)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsMeetingResponse, IsMeetingResponse);
        }

        if (IsNonDeliveryReport)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsNDR, IsNonDeliveryReport);
        }

        if (IsPermissionControlled)
        {
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.IsPermissionControlled,
                IsPermissionControlled
            );
        }

        if (_isReadReceipt)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsReadReceipt, IsReadReceipt);
        }

        if (IsSigned)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsSigned, IsSigned);
        }

        if (IsVoicemail)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsVoicemail, IsVoicemail);
        }

        if (ItemClasses.Count > 0)
        {
            ItemClasses.WriteToXml(writer, XmlElementNames.ItemClasses);
        }

        if (MessageClassifications.Count > 0)
        {
            MessageClassifications.WriteToXml(writer, XmlElementNames.MessageClassifications);
        }

        if (NotSentToMe)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.NotSentToMe, NotSentToMe);
        }

        if (SentCcMe)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SentCcMe, SentCcMe);
        }

        if (SentOnlyToMe)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SentOnlyToMe, SentOnlyToMe);
        }

        if (SentToAddresses.Count > 0)
        {
            SentToAddresses.WriteToXml(writer, XmlElementNames.SentToAddresses);
        }

        if (SentToMe)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SentToMe, SentToMe);
        }

        if (SentToOrCcMe)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SentToOrCcMe, SentToOrCcMe);
        }

        if (Sensitivity.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Sensitivity, Sensitivity.Value);
        }

        if (WithinDateRange.Start.HasValue || WithinDateRange.End.HasValue)
        {
            WithinDateRange.WriteToXml(writer, XmlElementNames.WithinDateRange);
        }

        if (WithinSizeRange.MaximumSize.HasValue || WithinSizeRange.MinimumSize.HasValue)
        {
            WithinSizeRange.WriteToXml(writer, XmlElementNames.WithinSizeRange);
        }
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void InternalValidate()
    {
        base.InternalValidate();

        EwsUtilities.ValidateParam(FromAddresses);
        EwsUtilities.ValidateParam(SentToAddresses);
        EwsUtilities.ValidateParam(WithinDateRange);
        EwsUtilities.ValidateParam(WithinSizeRange);
    }
}
