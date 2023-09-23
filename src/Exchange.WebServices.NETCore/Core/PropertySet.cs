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
///     Represents a set of item or folder properties. Property sets are used to indicate what properties of an item or
///     folder should be loaded when binding to an existing item or folder or when loading an item or folder's properties.
/// </summary>
[PublicAPI]
public sealed class PropertySet : ISelfValidate, IEnumerable<PropertyDefinitionBase>
{
    /// <summary>
    ///     Returns a predefined property set that only includes the Id property.
    /// </summary>
    public static readonly PropertySet IdOnly = CreateReadonlyPropertySet(BasePropertySet.IdOnly);

    /// <summary>
    ///     Returns a predefined property set that includes the first class properties of an item or folder.
    /// </summary>
    public static readonly PropertySet FirstClassProperties =
        CreateReadonlyPropertySet(BasePropertySet.FirstClassProperties);

    /// <summary>
    ///     Maps BasePropertySet values to EWS's BaseShape values.
    /// </summary>
    internal static readonly IReadOnlyDictionary<BasePropertySet, string> DefaultPropertySetMap =
        new Dictionary<BasePropertySet, string>
        { 
            // @formatter:off
            { BasePropertySet.IdOnly, "IdOnly" },
            { BasePropertySet.FirstClassProperties, "AllProperties" },
            // @formatter:on
        };

    /// <summary>
    ///     The list of additional properties included in this property set.
    /// </summary>
    private readonly List<PropertyDefinitionBase> _additionalProperties = new();

    /// <summary>
    ///     Value indicating whether or not to add a blank target attribute to anchor links.
    /// </summary>
    private bool? _addTargetToLinks;


    /// <summary>
    ///     The base property set this property set is based upon.
    /// </summary>
    private BasePropertySet _basePropertySet;

    /// <summary>
    ///     Value indicating whether or not the server should block references to external images.
    /// </summary>
    private bool? _blockExternalImages;

    /// <summary>
    ///     Value indicating whether or not the server should convert HTML code page to UTF8.
    /// </summary>
    private bool? _convertHtmlCodePageToUtf8;

    /// <summary>
    ///     Value indicating whether or not the server should filter HTML content.
    /// </summary>
    private bool? _filterHtml;

    /// <summary>
    ///     Value of the URL template to use for the src attribute of inline IMG elements.
    /// </summary>
    private string _inlineImageUrlTemplate;

    /// <summary>
    ///     Value indicating whether or not this PropertySet can be modified.
    /// </summary>
    private bool _isReadOnly;

    /// <summary>
    ///     Value indicating the maximum body size to retrieve.
    /// </summary>
    private int? _maximumBodySize;

    /// <summary>
    ///     The requested body type for get and find operations. If null, the "best body" is returned.
    /// </summary>
    private BodyType? _requestedBodyType;

    /// <summary>
    ///     The requested normalized body type for get and find operations. If null, the should return the same value as body
    ///     type.
    /// </summary>
    private BodyType? _requestedNormalizedBodyType;

    /// <summary>
    ///     The requested unique body type for get and find operations. If null, the should return the same value as body type.
    /// </summary>
    private BodyType? _requestedUniqueBodyType;

    /// <summary>
    ///     Gets or sets the base property set the property set is based upon.
    /// </summary>
    public BasePropertySet BasePropertySet
    {
        get => _basePropertySet;
        set
        {
            ThrowIfReadonly();
            _basePropertySet = value;
        }
    }

    /// <summary>
    ///     Gets or sets type of body that should be loaded on items. If RequestedBodyType is null, body is returned as HTML if
    ///     available, plain text otherwise.
    /// </summary>
    public BodyType? RequestedBodyType
    {
        get => _requestedBodyType;
        set
        {
            ThrowIfReadonly();
            _requestedBodyType = value;
        }
    }

    /// <summary>
    ///     Gets or sets type of body that should be loaded on items. If null, the should return the same value as body type.
    /// </summary>
    public BodyType? RequestedUniqueBodyType
    {
        get => _requestedUniqueBodyType;
        set
        {
            ThrowIfReadonly();
            _requestedUniqueBodyType = value;
        }
    }

    /// <summary>
    ///     Gets or sets type of normalized body that should be loaded on items. If null, the should return the same value as
    ///     body type.
    /// </summary>
    public BodyType? RequestedNormalizedBodyType
    {
        get => _requestedNormalizedBodyType;
        set
        {
            ThrowIfReadonly();
            _requestedNormalizedBodyType = value;
        }
    }

    /// <summary>
    ///     Gets the number of explicitly added properties in this set.
    /// </summary>
    public int Count => _additionalProperties.Count;

    /// <summary>
    ///     Gets or sets value indicating whether or not to filter potentially unsafe HTML content from message bodies.
    /// </summary>
    public bool? FilterHtmlContent
    {
        get => _filterHtml;
        set
        {
            ThrowIfReadonly();
            _filterHtml = value;
        }
    }

    /// <summary>
    ///     Gets or sets value indicating whether or not to convert HTML code page to UTF8 encoding.
    /// </summary>
    public bool? ConvertHtmlCodePageToUTF8
    {
        get => _convertHtmlCodePageToUtf8;
        set
        {
            ThrowIfReadonly();
            _convertHtmlCodePageToUtf8 = value;
        }
    }

    /// <summary>
    ///     Gets or sets a value of the URL template to use for the src attribute of inline IMG elements.
    /// </summary>
    public string InlineImageUrlTemplate
    {
        get => _inlineImageUrlTemplate;
        set
        {
            ThrowIfReadonly();
            _inlineImageUrlTemplate = value;
        }
    }

    /// <summary>
    ///     Gets or sets value indicating whether or not to convert inline images to data URLs.
    /// </summary>
    public bool? BlockExternalImages
    {
        get => _blockExternalImages;
        set
        {
            ThrowIfReadonly();
            _blockExternalImages = value;
        }
    }

    /// <summary>
    ///     Gets or sets value indicating whether or not to add blank target attribute to anchor links.
    /// </summary>
    public bool? AddBlankTargetToLinks
    {
        get => _addTargetToLinks;
        set
        {
            ThrowIfReadonly();
            _addTargetToLinks = value;
        }
    }

    /// <summary>
    ///     Gets or sets the maximum size of the body to be retrieved.
    /// </summary>
    /// <value>
    ///     The maximum size of the body to be retrieved.
    /// </value>
    public int? MaximumBodySize
    {
        get => _maximumBodySize;
        set
        {
            ThrowIfReadonly();
            _maximumBodySize = value;
        }
    }

    /// <summary>
    ///     Gets the <see cref="Microsoft.Exchange.WebServices.Data.PropertyDefinitionBase" /> at the specified index.
    /// </summary>
    /// <param name="index">Index.</param>
    public PropertyDefinitionBase this[int index] => _additionalProperties[index];

    /// <summary>
    ///     Initializes a new instance of PropertySet.
    /// </summary>
    /// <param name="basePropertySet">The base property set to base the property set upon.</param>
    /// <param name="additionalProperties">
    ///     Additional properties to include in the property set. Property definitions are
    ///     available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start,
    ///     ContactSchema.GivenName, etc.)
    /// </param>
    public PropertySet(BasePropertySet basePropertySet, params PropertyDefinitionBase[]? additionalProperties)
        : this(basePropertySet, (IEnumerable<PropertyDefinitionBase>?)additionalProperties)
    {
    }

    /// <summary>
    ///     Initializes a new instance of PropertySet.
    /// </summary>
    /// <param name="basePropertySet">The base property set to base the property set upon.</param>
    /// <param name="additionalProperties">
    ///     Additional properties to include in the property set. Property definitions are
    ///     available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start,
    ///     ContactSchema.GivenName, etc.)
    /// </param>
    public PropertySet(BasePropertySet basePropertySet, IEnumerable<PropertyDefinitionBase>? additionalProperties)
    {
        _basePropertySet = basePropertySet;

        if (additionalProperties != null)
        {
            _additionalProperties.AddRange(additionalProperties);
        }
    }

    /// <summary>
    ///     Initializes a new instance of PropertySet based upon BasePropertySet.IdOnly.
    /// </summary>
    public PropertySet()
        : this(BasePropertySet.IdOnly, null)
    {
    }

    /// <summary>
    ///     Initializes a new instance of PropertySet.
    /// </summary>
    /// <param name="basePropertySet">The base property set to base the property set upon.</param>
    public PropertySet(BasePropertySet basePropertySet)
        : this(basePropertySet, null)
    {
    }

    /// <summary>
    ///     Initializes a new instance of PropertySet based upon BasePropertySet.IdOnly.
    /// </summary>
    /// <param name="additionalProperties">
    ///     Additional properties to include in the property set. Property definitions are
    ///     available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start,
    ///     ContactSchema.GivenName, etc.)
    /// </param>
    public PropertySet(params PropertyDefinitionBase[] additionalProperties)
        : this(BasePropertySet.IdOnly, additionalProperties)
    {
    }

    /// <summary>
    ///     Initializes a new instance of PropertySet based upon BasePropertySet.IdOnly.
    /// </summary>
    /// <param name="additionalProperties">
    ///     Additional properties to include in the property set. Property definitions are
    ///     available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start,
    ///     ContactSchema.GivenName, etc.)
    /// </param>
    public PropertySet(IEnumerable<PropertyDefinitionBase> additionalProperties)
        : this(BasePropertySet.IdOnly, additionalProperties)
    {
    }


    #region IEnumerable<PropertyDefinitionBase> Members

    /// <summary>
    ///     Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>
    ///     A <see cref="T:System.Collections.Generic.IEnumerator`1" /> that can be used to iterate through the collection.
    /// </returns>
    public IEnumerator<PropertyDefinitionBase> GetEnumerator()
    {
        return _additionalProperties.GetEnumerator();
    }

    #endregion


    #region IEnumerable Members

    /// <summary>
    ///     Returns an enumerator that iterates through a collection.
    /// </summary>
    /// <returns>
    ///     An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.
    /// </returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return _additionalProperties.GetEnumerator();
    }

    #endregion


    /// <summary>
    ///     Implements ISelfValidate.Validate. Validates this property set.
    /// </summary>
    void ISelfValidate.Validate()
    {
        InternalValidate();
    }

    /// <summary>
    ///     Implements an implicit conversion between PropertySet and BasePropertySet.
    /// </summary>
    /// <param name="basePropertySet">The BasePropertySet value to convert from.</param>
    /// <returns>A PropertySet instance based on the specified base property set.</returns>
    public static implicit operator PropertySet(BasePropertySet basePropertySet)
    {
        return new PropertySet(basePropertySet);
    }

    /// <summary>
    ///     Adds the specified property to the property set.
    /// </summary>
    /// <param name="property">The property to add.</param>
    public void Add(PropertyDefinitionBase property)
    {
        ThrowIfReadonly();
        EwsUtilities.ValidateParam(property);

        if (!_additionalProperties.Contains(property))
        {
            _additionalProperties.Add(property);
        }
    }

    /// <summary>
    ///     Adds the specified properties to the property set.
    /// </summary>
    /// <param name="properties">The properties to add.</param>
    public void AddRange(IEnumerable<PropertyDefinitionBase> properties)
    {
        ThrowIfReadonly();
        EwsUtilities.ValidateParamCollection(properties);

        foreach (var property in properties)
        {
            Add(property);
        }
    }

    /// <summary>
    ///     Remove all explicitly added properties from the property set.
    /// </summary>
    public void Clear()
    {
        ThrowIfReadonly();
        _additionalProperties.Clear();
    }

    /// <summary>
    ///     Creates a read-only PropertySet.
    /// </summary>
    /// <param name="basePropertySet">The base property set.</param>
    /// <returns>PropertySet</returns>
    private static PropertySet CreateReadonlyPropertySet(BasePropertySet basePropertySet)
    {
        return new PropertySet(basePropertySet)
        {
            _isReadOnly = true,
        };
    }

    /// <summary>
    ///     Gets the name of the shape.
    /// </summary>
    /// <param name="serviceObjectType">Type of the service object.</param>
    /// <returns>Shape name.</returns>
    private static string GetShapeName(ServiceObjectType serviceObjectType)
    {
        switch (serviceObjectType)
        {
            case ServiceObjectType.Item:
            {
                return XmlElementNames.ItemShape;
            }
            case ServiceObjectType.Folder:
            {
                return XmlElementNames.FolderShape;
            }
            case ServiceObjectType.Conversation:
            {
                return XmlElementNames.ConversationShape;
            }
            case ServiceObjectType.Persona:
            {
                return XmlElementNames.PersonaShape;
            }
            default:
            {
                EwsUtilities.Assert(
                    false,
                    "PropertySet.GetShapeName",
                    string.Format(
                        "An unexpected object type {0} for property shape. This code path should never be reached.",
                        serviceObjectType
                    )
                );
                return string.Empty;
            }
        }
    }

    /// <summary>
    ///     Throws if readonly property set.
    /// </summary>
    private void ThrowIfReadonly()
    {
        if (_isReadOnly)
        {
            throw new NotSupportedException(Strings.PropertySetCannotBeModified);
        }
    }

    /// <summary>
    ///     Determines whether the specified property has been explicitly added to this property set using the Add or AddRange
    ///     methods.
    /// </summary>
    /// <param name="property">The property.</param>
    /// <returns>
    ///     <c>true</c> if this property set contains the specified property; otherwise, <c>false</c>.
    /// </returns>
    public bool Contains(PropertyDefinitionBase property)
    {
        return _additionalProperties.Contains(property);
    }

    /// <summary>
    ///     Removes the specified property from the set.
    /// </summary>
    /// <param name="property">The property to remove.</param>
    /// <returns>true if the property was successfully removed, false otherwise.</returns>
    public bool Remove(PropertyDefinitionBase property)
    {
        ThrowIfReadonly();
        return _additionalProperties.Remove(property);
    }


    /// <summary>
    ///     Writes additional properties to XML.
    /// </summary>
    /// <param name="writer">The writer to write to.</param>
    /// <param name="propertyDefinitions">The property definitions to write.</param>
    internal static void WriteAdditionalPropertiesToXml(
        EwsServiceXmlWriter writer,
        IEnumerable<PropertyDefinitionBase> propertyDefinitions
    )
    {
        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.AdditionalProperties);

        foreach (var propertyDefinition in propertyDefinitions)
        {
            propertyDefinition.WriteToXml(writer);
        }

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Validates this property set.
    /// </summary>
    internal void InternalValidate()
    {
        for (var i = 0; i < _additionalProperties.Count; i++)
        {
            if (_additionalProperties[i] == null)
            {
                throw new ServiceValidationException(string.Format(Strings.AdditionalPropertyIsNull, i));
            }
        }
    }

    /// <summary>
    ///     Validates this property set instance for request to ensure that:
    ///     1. Properties are valid for the request server version.
    ///     2. If only summary properties are legal for this request (e.g. FindItem) then only summary properties were
    ///     specified.
    /// </summary>
    /// <param name="request">The request.</param>
    /// <param name="summaryPropertiesOnly">if set to <c>true</c> then only summary properties are allowed.</param>
    internal void ValidateForRequest(ServiceRequestBase request, bool summaryPropertiesOnly)
    {
        foreach (var propDefBase in _additionalProperties)
        {
            if (propDefBase is PropertyDefinition propertyDefinition)
            {
                if (propertyDefinition.Version > request.Service.RequestedServerVersion)
                {
                    throw new ServiceVersionException(
                        string.Format(
                            Strings.PropertyIncompatibleWithRequestVersion,
                            propertyDefinition.Name,
                            propertyDefinition.Version
                        )
                    );
                }

                if (summaryPropertiesOnly &&
                    !propertyDefinition.HasFlag(
                        PropertyDefinitionFlags.CanFind,
                        request.Service.RequestedServerVersion
                    ))
                {
                    throw new ServiceValidationException(
                        string.Format(
                            Strings.NonSummaryPropertyCannotBeUsed,
                            propertyDefinition.Name,
                            request.GetXmlElementName()
                        )
                    );
                }
            }
        }

        if (FilterHtmlContent.HasValue)
        {
            if (request.Service.RequestedServerVersion < ExchangeVersion.Exchange2010)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PropertyIncompatibleWithRequestVersion,
                        "FilterHtmlContent",
                        ExchangeVersion.Exchange2010
                    )
                );
            }
        }

        if (ConvertHtmlCodePageToUTF8.HasValue)
        {
            if (request.Service.RequestedServerVersion < ExchangeVersion.Exchange2010_SP1)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PropertyIncompatibleWithRequestVersion,
                        "ConvertHtmlCodePageToUTF8",
                        ExchangeVersion.Exchange2010_SP1
                    )
                );
            }
        }

        if (!string.IsNullOrEmpty(InlineImageUrlTemplate))
        {
            if (request.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PropertyIncompatibleWithRequestVersion,
                        "InlineImageUrlTemplate",
                        ExchangeVersion.Exchange2013
                    )
                );
            }
        }

        if (BlockExternalImages.HasValue)
        {
            if (request.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PropertyIncompatibleWithRequestVersion,
                        "BlockExternalImages",
                        ExchangeVersion.Exchange2013
                    )
                );
            }
        }

        if (AddBlankTargetToLinks.HasValue)
        {
            if (request.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PropertyIncompatibleWithRequestVersion,
                        "AddTargetToLinks",
                        ExchangeVersion.Exchange2013
                    )
                );
            }
        }

        if (MaximumBodySize.HasValue)
        {
            if (request.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PropertyIncompatibleWithRequestVersion,
                        "MaximumBodySize",
                        ExchangeVersion.Exchange2013
                    )
                );
            }
        }
    }

    /// <summary>
    ///     Writes the property set to XML.
    /// </summary>
    /// <param name="writer">The writer to write to.</param>
    /// <param name="serviceObjectType">The type of service object the property set is emitted for.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer, ServiceObjectType serviceObjectType)
    {
        var shapeElementName = GetShapeName(serviceObjectType);

        writer.WriteStartElement(XmlNamespace.Messages, shapeElementName);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.BaseShape, DefaultPropertySetMap[BasePropertySet]);

        if (serviceObjectType == ServiceObjectType.Item)
        {
            if (RequestedBodyType.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.BodyType, RequestedBodyType.Value);
            }

            if (RequestedUniqueBodyType.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.UniqueBodyType,
                    RequestedUniqueBodyType.Value
                );
            }

            if (RequestedNormalizedBodyType.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.NormalizedBodyType,
                    RequestedNormalizedBodyType.Value
                );
            }

            if (FilterHtmlContent.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.FilterHtmlContent,
                    FilterHtmlContent.Value
                );
            }

            if (ConvertHtmlCodePageToUTF8.HasValue &&
                writer.Service.RequestedServerVersion >= ExchangeVersion.Exchange2010_SP1)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.ConvertHtmlCodePageToUTF8,
                    ConvertHtmlCodePageToUTF8.Value
                );
            }

            if (!string.IsNullOrEmpty(InlineImageUrlTemplate) &&
                writer.Service.RequestedServerVersion >= ExchangeVersion.Exchange2013)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.InlineImageUrlTemplate,
                    InlineImageUrlTemplate
                );
            }

            if (BlockExternalImages.HasValue && writer.Service.RequestedServerVersion >= ExchangeVersion.Exchange2013)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.BlockExternalImages,
                    BlockExternalImages.Value
                );
            }

            if (AddBlankTargetToLinks.HasValue && writer.Service.RequestedServerVersion >= ExchangeVersion.Exchange2013)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.AddBlankTargetToLinks,
                    AddBlankTargetToLinks.Value
                );
            }

            if (MaximumBodySize.HasValue && writer.Service.RequestedServerVersion >= ExchangeVersion.Exchange2013)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MaximumBodySize, MaximumBodySize.Value);
            }
        }

        if (_additionalProperties.Count > 0)
        {
            WriteAdditionalPropertiesToXml(writer, _additionalProperties);
        }

        writer.WriteEndElement(); // Item/FolderShape
    }
}
