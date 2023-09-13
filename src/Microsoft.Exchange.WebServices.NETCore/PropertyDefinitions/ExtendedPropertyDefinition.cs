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

using System.Text;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the definition of an extended property.
/// </summary>
[PublicAPI]
public sealed class ExtendedPropertyDefinition : PropertyDefinitionBase
{
    #region Constants

    private const string FieldFormat = "{0}: {1} ";

    private const string PropertySetFieldName = nameof(PropertySet);
    private const string PropertySetIdFieldName = nameof(PropertySetId);
    private const string TagFieldName = nameof(Tag);
    private const string NameFieldName = nameof(Name);
    private const string IdFieldName = nameof(Id);
    private const string MapiTypeFieldName = nameof(MapiType);

    #endregion


    #region Fields

    private DefaultExtendedPropertySet? _propertySet;
    private Guid? _propertySetId;
    private int? _tag;
    private int? _id;

    #endregion


    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtendedPropertyDefinition" /> class.
    /// </summary>
    internal ExtendedPropertyDefinition()
    {
        MapiType = MapiPropertyType.String;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtendedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="mapiType">The MAPI type of the extended property.</param>
    internal ExtendedPropertyDefinition(MapiPropertyType mapiType)
        : this()
    {
        MapiType = mapiType;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtendedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="tag">The tag of the extended property.</param>
    /// <param name="mapiType">The MAPI type of the extended property.</param>
    public ExtendedPropertyDefinition(int tag, MapiPropertyType mapiType)
        : this(mapiType)
    {
        if (tag < 0 || tag > ushort.MaxValue)
        {
            throw new ArgumentOutOfRangeException(nameof(tag), Strings.TagValueIsOutOfRange);
        }

        _tag = tag;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtendedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="propertySet">The extended property set of the extended property.</param>
    /// <param name="name">The name of the extended property.</param>
    /// <param name="mapiType">The MAPI type of the extended property.</param>
    public ExtendedPropertyDefinition(DefaultExtendedPropertySet propertySet, string name, MapiPropertyType mapiType)
        : this(mapiType)
    {
        EwsUtilities.ValidateParam(name, nameof(name));

        _propertySet = propertySet;
        Name = name;
    }

    /// <summary>
    ///     Initializes a new instance of ExtendedPropertyDefinition.
    /// </summary>
    /// <param name="propertySet">The property set of the extended property.</param>
    /// <param name="id">The Id of the extended property.</param>
    /// <param name="mapiType">The MAPI type of the extended property.</param>
    public ExtendedPropertyDefinition(DefaultExtendedPropertySet propertySet, int id, MapiPropertyType mapiType)
        : this(mapiType)
    {
        _propertySet = propertySet;
        _id = id;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtendedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="propertySetId">The property set Id of the extended property.</param>
    /// <param name="name">The name of the extended property.</param>
    /// <param name="mapiType">The MAPI type of the extended property.</param>
    public ExtendedPropertyDefinition(Guid propertySetId, string name, MapiPropertyType mapiType)
        : this(mapiType)
    {
        EwsUtilities.ValidateParam(name, nameof(name));

        _propertySetId = propertySetId;
        Name = name;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtendedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="propertySetId">The property set Id of the extended property.</param>
    /// <param name="id">The Id of the extended property.</param>
    /// <param name="mapiType">The MAPI type of the extended property.</param>
    public ExtendedPropertyDefinition(Guid propertySetId, int id, MapiPropertyType mapiType)
        : this(mapiType)
    {
        _propertySetId = propertySetId;
        _id = id;
    }

    /// <summary>
    ///     Determines whether two specified instances of ExtendedPropertyDefinition are equal.
    /// </summary>
    /// <param name="extPropDef1">First extended property definition.</param>
    /// <param name="extPropDef2">Second extended property definition.</param>
    /// <returns>True if extended property definitions are equal.</returns>
    internal static bool IsEqualTo(ExtendedPropertyDefinition? extPropDef1, ExtendedPropertyDefinition? extPropDef2)
    {
        return ReferenceEquals(extPropDef1, extPropDef2) ||
               (extPropDef1 is not null &&
                extPropDef2 is not null &&
                extPropDef1.Id == extPropDef2.Id &&
                extPropDef1.MapiType == extPropDef2.MapiType &&
                extPropDef1.Tag == extPropDef2.Tag &&
                extPropDef1.Name == extPropDef2.Name &&
                extPropDef1.PropertySet == extPropDef2.PropertySet &&
                extPropDef1._propertySetId == extPropDef2._propertySetId);
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.ExtendedFieldURI;
    }

    /// <summary>
    ///     Gets the minimum Exchange version that supports this extended property.
    /// </summary>
    /// <value>The version.</value>
    public override ExchangeVersion Version => ExchangeVersion.Exchange2007_SP1;

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        if (_propertySet.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.DistinguishedPropertySetId, _propertySet.Value);
        }

        if (_propertySetId.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.PropertySetId, _propertySetId.Value.ToString());
        }

        if (_tag.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.PropertyTag, _tag.Value);
        }

        if (!string.IsNullOrEmpty(Name))
        {
            writer.WriteAttributeValue(XmlAttributeNames.PropertyName, Name);
        }

        if (_id.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.PropertyId, _id.Value);
        }

        writer.WriteAttributeValue(XmlAttributeNames.PropertyType, MapiType);
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        var attributeValue = reader.ReadAttributeValue(XmlAttributeNames.DistinguishedPropertySetId);
        if (!string.IsNullOrEmpty(attributeValue))
        {
            _propertySet = (DefaultExtendedPropertySet)Enum.Parse(
                typeof(DefaultExtendedPropertySet),
                attributeValue,
                false
            );
        }

        attributeValue = reader.ReadAttributeValue(XmlAttributeNames.PropertySetId);
        if (!string.IsNullOrEmpty(attributeValue))
        {
            _propertySetId = new Guid(attributeValue);
        }

        attributeValue = reader.ReadAttributeValue(XmlAttributeNames.PropertyTag);
        if (!string.IsNullOrEmpty(attributeValue))
        {
            _tag = Convert.ToUInt16(attributeValue, 16);
        }

        Name = reader.ReadAttributeValue(XmlAttributeNames.PropertyName);

        attributeValue = reader.ReadAttributeValue(XmlAttributeNames.PropertyId);
        if (!string.IsNullOrEmpty(attributeValue))
        {
            _id = int.Parse(attributeValue);
        }

        MapiType = reader.ReadAttributeValue<MapiPropertyType>(XmlAttributeNames.PropertyType);
    }

    /// <summary>
    ///     Determines whether two specified instances of ExtendedPropertyDefinition are equal.
    /// </summary>
    /// <param name="extPropDef1">First extended property definition.</param>
    /// <param name="extPropDef2">Second extended property definition.</param>
    /// <returns>True if extended property definitions are equal.</returns>
    public static bool operator ==(ExtendedPropertyDefinition? extPropDef1, ExtendedPropertyDefinition? extPropDef2)
    {
        return IsEqualTo(extPropDef1, extPropDef2);
    }

    /// <summary>
    ///     Determines whether two specified instances of ExtendedPropertyDefinition are not equal.
    /// </summary>
    /// <param name="extPropDef1">First extended property definition.</param>
    /// <param name="extPropDef2">Second extended property definition.</param>
    /// <returns>True if extended property definitions are equal.</returns>
    public static bool operator !=(ExtendedPropertyDefinition? extPropDef1, ExtendedPropertyDefinition? extPropDef2)
    {
        return !IsEqualTo(extPropDef1, extPropDef2);
    }

    /// <summary>
    ///     Determines whether a given extended property definition is equal to this extended property definition.
    /// </summary>
    /// <param name="obj">The object to check for equality.</param>
    /// <returns>True if the properties definitions define the same extended property.</returns>
    public override bool Equals(object? obj)
    {
        var propertyDefinition = obj as ExtendedPropertyDefinition;
        return IsEqualTo(propertyDefinition, this);
    }

    /// <summary>
    ///     Serves as a hash function for a particular type.
    /// </summary>
    /// <returns>
    ///     A hash code for the current <see cref="T:System.Object" />.
    /// </returns>
    public override int GetHashCode()
    {
        return GetPrintableName().GetHashCode();
    }

    /// <summary>
    ///     Gets the property definition's printable name.
    /// </summary>
    /// <returns>
    ///     The property definition's printable name.
    /// </returns>
    internal override string GetPrintableName()
    {
        var sb = new StringBuilder();
        sb.Append('{');
        sb.Append(FormatField(NameFieldName, Name));
        sb.Append(FormatField<MapiPropertyType?>(MapiTypeFieldName, MapiType));
        sb.Append(FormatField(IdFieldName, Id));
        sb.Append(FormatField(PropertySetFieldName, PropertySet));
        sb.Append(FormatField(PropertySetIdFieldName, PropertySetId));
        sb.Append(FormatField(TagFieldName, Tag));
        sb.Append('}');
        return sb.ToString();
    }

    /// <summary>
    ///     Formats the field.
    /// </summary>
    /// <typeparam name="T">Type of field value.</typeparam>
    /// <param name="name">The name.</param>
    /// <param name="fieldValue">The field value.</param>
    /// <returns>Formatted value.</returns>
    internal static string FormatField<T>(string name, T fieldValue)
    {
        return fieldValue != null ? string.Format(FieldFormat, name, fieldValue.ToString()) : string.Empty;
    }

    /// <summary>
    ///     Gets the property set of the extended property.
    /// </summary>
    public DefaultExtendedPropertySet? PropertySet => _propertySet;

    /// <summary>
    ///     Gets the property set Id or the extended property.
    /// </summary>
    public Guid? PropertySetId => _propertySetId;

    /// <summary>
    ///     Gets the extended property's tag.
    /// </summary>
    public int? Tag => _tag;

    /// <summary>
    ///     Gets the name of the extended property.
    /// </summary>
    public string? Name { get; private set; }

    /// <summary>
    ///     Gets the Id of the extended property.
    /// </summary>
    public int? Id => _id;

    /// <summary>
    ///     Gets the MAPI type of the extended property.
    /// </summary>
    public MapiPropertyType MapiType { get; private set; }

    /// <summary>
    ///     Gets the property type.
    /// </summary>
    public override Type Type => MapiTypeConverter.MapiTypeConverterMap[MapiType].Type;
}
