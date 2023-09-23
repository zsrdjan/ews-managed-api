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
///     Represents an attachment to an item.
/// </summary>
[PublicAPI]
public abstract class Attachment : ComplexProperty
{
    private string _contentId;
    private string _contentLocation;
    private string _contentType;
    private bool _isInline;
    private DateTime _lastModifiedTime;
    private string _name;
    private int _size;

    /// <summary>
    ///     Gets the Id of the attachment.
    /// </summary>
    public string Id { get; internal set; }

    /// <summary>
    ///     Gets or sets the name of the attachment.
    /// </summary>
    public string Name
    {
        get => _name;
        set => SetFieldValue(ref _name, value);
    }

    /// <summary>
    ///     Gets or sets the content type of the attachment.
    /// </summary>
    public string ContentType
    {
        get => _contentType;
        set => SetFieldValue(ref _contentType, value);
    }

    /// <summary>
    ///     Gets or sets the content Id of the attachment. ContentId can be used as a custom way to identify
    ///     an attachment in order to reference it from within the body of the item the attachment belongs to.
    /// </summary>
    public string ContentId
    {
        get => _contentId;
        set => SetFieldValue(ref _contentId, value);
    }

    /// <summary>
    ///     Gets or sets the content location of the attachment. ContentLocation can be used to associate
    ///     an attachment with a Url defining its location on the Web.
    /// </summary>
    public string ContentLocation
    {
        get => _contentLocation;
        set => SetFieldValue(ref _contentLocation, value);
    }

    /// <summary>
    ///     Gets the size of the attachment.
    /// </summary>
    public int Size
    {
        get
        {
            EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, nameof(Size));

            return _size;
        }

        internal set
        {
            EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, nameof(Size));

            SetFieldValue(ref _size, value);
        }
    }

    /// <summary>
    ///     Gets the date and time when this attachment was last modified.
    /// </summary>
    public DateTime LastModifiedTime
    {
        get
        {
            EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, nameof(LastModifiedTime));

            return _lastModifiedTime;
        }

        internal set
        {
            EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, nameof(LastModifiedTime));

            SetFieldValue(ref _lastModifiedTime, value);
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether this is an inline attachment.
    ///     Inline attachments are not visible to end users.
    /// </summary>
    public bool IsInline
    {
        get
        {
            EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, nameof(IsInline));

            return _isInline;
        }

        set
        {
            EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, nameof(IsInline));

            SetFieldValue(ref _isInline, value);
        }
    }

    /// <summary>
    ///     True if the attachment has not yet been saved, false otherwise.
    /// </summary>
    internal bool IsNew => string.IsNullOrEmpty(Id);

    /// <summary>
    ///     Gets the owner of the attachment.
    /// </summary>
    internal Item? Owner { get; }

    /// <summary>
    ///     Gets the related exchange service.
    /// </summary>
    internal ExchangeService Service { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Attachment" /> class.
    /// </summary>
    /// <param name="owner">The owner.</param>
    internal Attachment(Item? owner)
    {
        Owner = owner;

        if (Owner != null)
        {
            Service = Owner.Service;
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Attachment" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal Attachment(ExchangeService service)
    {
        Service = service;
    }

    /// <summary>
    ///     Throws exception if this is not a new service object.
    /// </summary>
    internal void ThrowIfThisIsNotNew()
    {
        if (!IsNew)
        {
            throw new InvalidOperationException(Strings.AttachmentCannotBeUpdated);
        }
    }

    /// <summary>
    ///     Sets value of field.
    /// </summary>
    /// <remarks>
    ///     We override the base implementation. Attachments cannot be modified so any attempts
    ///     the change a property on an existing attachment is an error.
    /// </remarks>
    /// <typeparam name="T">Field type.</typeparam>
    /// <param name="field">The field.</param>
    /// <param name="value">The value.</param>
    internal override void SetFieldValue<T>(ref T field, T value)
    {
        ThrowIfThisIsNotNew();
        base.SetFieldValue(ref field, value);
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal abstract string GetXmlElementName();

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.AttachmentId:
            {
                Id = reader.ReadAttributeValue(XmlAttributeNames.Id);

                if (Owner != null)
                {
                    var rootItemChangeKey = reader.ReadAttributeValue(XmlAttributeNames.RootItemChangeKey);

                    if (!string.IsNullOrEmpty(rootItemChangeKey))
                    {
                        Owner.RootItemId.ChangeKey = rootItemChangeKey;
                    }
                }

                reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.AttachmentId);
                return true;
            }
            case XmlElementNames.Name:
            {
                _name = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.ContentType:
            {
                _contentType = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.ContentId:
            {
                _contentId = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.ContentLocation:
            {
                _contentLocation = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.Size:
            {
                _size = reader.ReadElementValue<int>();
                return true;
            }
            case XmlElementNames.LastModifiedTime:
            {
                _lastModifiedTime = reader.ReadElementValueAsDateTime().Value;
                return true;
            }
            case XmlElementNames.IsInline:
            {
                _isInline = reader.ReadElementValue<bool>();
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
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentType, ContentType);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentId, ContentId);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentLocation, ContentLocation);

        if (writer.Service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsInline, IsInline);
        }
    }

    /// <summary>
    ///     Load the attachment.
    /// </summary>
    /// <param name="bodyType">Type of the body.</param>
    /// <param name="additionalProperties">The additional properties.</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<GetAttachmentResponse>> InternalLoad(
        BodyType? bodyType,
        IEnumerable<PropertyDefinitionBase>? additionalProperties,
        CancellationToken token
    )
    {
        return Service.GetAttachment(this, bodyType, additionalProperties, token);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    /// <param name="attachmentIndex">Index of this attachment.</param>
    internal virtual void Validate(int attachmentIndex)
    {
    }

    /// <summary>
    ///     Loads the attachment. Calling this method results in a call to EWS.
    /// </summary>
    public Task<ServiceResponseCollection<GetAttachmentResponse>> Load(CancellationToken token = default)
    {
        return InternalLoad(null, null, token);
    }
}
