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
///     Represents an e-mail address.
/// </summary>
[PublicAPI]
public class EmailAddress : ComplexProperty, ISearchStringProvider
{
    /// <summary>
    ///     SMTP routing type.
    /// </summary>
    internal const string SmtpRoutingType = "SMTP";

    /// <summary>
    ///     Email address.
    /// </summary>
    private string _address;

    /// <summary>
    ///     ItemId - Contact or PDL.
    /// </summary>
    private ItemId? _id;

    /// <summary>
    ///     Mailbox type.
    /// </summary>
    private MailboxType? _mailboxType;

    /// <summary>
    ///     Display name.
    /// </summary>
    private string _name;

    /// <summary>
    ///     Routing type.
    /// </summary>
    private string _routingType;

    /// <summary>
    ///     Gets or sets the name associated with the e-mail address.
    /// </summary>
    public string Name
    {
        get => _name;
        set => SetFieldValue(ref _name, value);
    }

    /// <summary>
    ///     Gets or sets the actual address associated with the e-mail address. The type of the Address property
    ///     must match the specified routing type. If RoutingType is not set, Address is assumed to be an SMTP
    ///     address.
    /// </summary>
    public string Address
    {
        get => _address;
        set => SetFieldValue(ref _address, value);
    }

    /// <summary>
    ///     Gets or sets the routing type associated with the e-mail address. If RoutingType is not set,
    ///     Address is assumed to be an SMTP address.
    /// </summary>
    public string RoutingType
    {
        get => _routingType;
        set => SetFieldValue(ref _routingType, value);
    }

    /// <summary>
    ///     Gets or sets the type of the e-mail address.
    /// </summary>
    public MailboxType? MailboxType
    {
        get => _mailboxType;
        set => SetFieldValue(ref _mailboxType, value);
    }

    /// <summary>
    ///     Gets or sets the Id of the contact the e-mail address represents. When Id is specified, Address
    ///     should be set to null.
    /// </summary>
    public ItemId? Id
    {
        get => _id;
        set => SetFieldValue(ref _id, value);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddress" /> class.
    /// </summary>
    public EmailAddress()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddress" /> class.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address used to initialize the EmailAddress.</param>
    public EmailAddress(string smtpAddress)
        : this()
    {
        _address = smtpAddress;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddress" /> class.
    /// </summary>
    /// <param name="name">The name used to initialize the EmailAddress.</param>
    /// <param name="smtpAddress">The SMTP address used to initialize the EmailAddress.</param>
    public EmailAddress(string name, string smtpAddress)
        : this(smtpAddress)
    {
        _name = name;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddress" /> class.
    /// </summary>
    /// <param name="name">The name used to initialize the EmailAddress.</param>
    /// <param name="address">The address used to initialize the EmailAddress.</param>
    /// <param name="routingType">The routing type used to initialize the EmailAddress.</param>
    public EmailAddress(string name, string address, string routingType)
        : this(name, address)
    {
        _routingType = routingType;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddress" /> class.
    /// </summary>
    /// <param name="name">The name used to initialize the EmailAddress.</param>
    /// <param name="address">The address used to initialize the EmailAddress.</param>
    /// <param name="routingType">The routing type used to initialize the EmailAddress.</param>
    /// <param name="mailboxType">Mailbox type of the participant.</param>
    internal EmailAddress(string name, string address, string routingType, MailboxType mailboxType)
        : this(name, address, routingType)
    {
        _mailboxType = mailboxType;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddress" /> class.
    /// </summary>
    /// <param name="name">The name used to initialize the EmailAddress.</param>
    /// <param name="address">The address used to initialize the EmailAddress.</param>
    /// <param name="routingType">The routing type used to initialize the EmailAddress.</param>
    /// <param name="mailboxType">Mailbox type of the participant.</param>
    /// <param name="itemId">ItemId of a Contact or PDL.</param>
    internal EmailAddress(string? name, string address, string routingType, MailboxType mailboxType, ItemId itemId)
        : this(name, address, routingType)
    {
        _mailboxType = mailboxType;
        _id = itemId;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddress" /> class from another EmailAddress instance.
    /// </summary>
    /// <param name="mailbox">EMailAddress instance to copy.</param>
    internal EmailAddress(EmailAddress mailbox)
        : this()
    {
        EwsUtilities.ValidateParam(mailbox);

        Name = mailbox.Name;
        Address = mailbox.Address;
        RoutingType = mailbox.RoutingType;
        MailboxType = mailbox.MailboxType;
        Id = mailbox.Id;
    }


    #region ISearchStringProvider methods

    /// <summary>
    ///     Get a string representation for using this instance in a search filter.
    /// </summary>
    /// <returns>String representation of instance.</returns>
    string ISearchStringProvider.GetSearchString()
    {
        return Address;
    }

    #endregion


    /// <summary>
    ///     Defines an implicit conversion between a string representing an SMTP address and EmailAddress.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address to convert to EmailAddress.</param>
    /// <returns>An EmailAddress initialized with the specified SMTP address.</returns>
    public static implicit operator EmailAddress(string smtpAddress)
    {
        return new EmailAddress(smtpAddress);
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
            case XmlElementNames.Name:
            {
                _name = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.EmailAddress:
            {
                _address = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.RoutingType:
            {
                _routingType = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.MailboxType:
            {
                _mailboxType = reader.ReadElementValue<MailboxType>();
                return true;
            }
            case XmlElementNames.ItemId:
            {
                _id = new ItemId();
                _id.LoadFromXml(reader, reader.LocalName);
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
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Name, Name);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EmailAddress, Address);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RoutingType, RoutingType);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MailboxType, MailboxType);

        if (Id != null)
        {
            Id.WriteToXml(writer, XmlElementNames.ItemId);
        }
    }


    #region Object method overrides

    /// <summary>
    ///     Returns a <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </summary>
    /// <returns>
    ///     A <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </returns>
    public override string ToString()
    {
        string addressPart;

        if (string.IsNullOrEmpty(Address))
        {
            return string.Empty;
        }

        if (!string.IsNullOrEmpty(RoutingType))
        {
            addressPart = RoutingType + ":" + Address;
        }
        else
        {
            addressPart = Address;
        }

        if (!string.IsNullOrEmpty(Name))
        {
            return Name + " <" + addressPart + ">";
        }

        return addressPart;
    }

    #endregion
}
