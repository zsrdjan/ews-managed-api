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
///     Represents the retention tag of an item.
/// </summary>
[PublicAPI]
public class RetentionTagBase : ComplexProperty
{
    /// <summary>
    ///     Xml element name.
    /// </summary>
    private readonly string _xmlElementName;

    /// <summary>
    ///     Is explicit.
    /// </summary>
    private bool _isExplicit;

    /// <summary>
    ///     Retention id.
    /// </summary>
    private Guid _retentionId;

    /// <summary>
    ///     Gets or sets if the tag is explicit.
    /// </summary>
    public bool IsExplicit
    {
        get => _isExplicit;
        set => SetFieldValue(ref _isExplicit, value);
    }

    /// <summary>
    ///     Gets or sets the retention id.
    /// </summary>
    public Guid RetentionId
    {
        get => _retentionId;
        set => SetFieldValue(ref _retentionId, value);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="RetentionTagBase" /> class.
    /// </summary>
    /// <param name="xmlElementName">Xml element name.</param>
    public RetentionTagBase(string xmlElementName)
    {
        _xmlElementName = xmlElementName;
    }

    /// <summary>
    ///     Reads attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        _isExplicit = reader.ReadAttributeValue<bool>(XmlAttributeNames.IsExplicit);
    }

    /// <summary>
    ///     Reads text value from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
    {
        _retentionId = new Guid(reader.ReadValue());
    }

    /// <summary>
    ///     Writes attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.IsExplicit, _isExplicit);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (_retentionId != Guid.Empty)
        {
            writer.WriteValue(_retentionId.ToString(), _xmlElementName);
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
        if (_retentionId == Guid.Empty)
        {
            return string.Empty;
        }

        return _retentionId.ToString();
    }

    #endregion
}
