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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents generic property definition.
/// </summary>
/// <typeparam name="TPropertyValue">Property value type. Constrained to be a value type.</typeparam>
internal class GenericPropertyDefinition<TPropertyValue> : TypedPropertyDefinition
    where TPropertyValue : struct
{
    /// <summary>
    ///     Gets the property type.
    /// </summary>
    public override Type Type => IsNullable ? typeof(TPropertyValue?) : typeof(TPropertyValue);

    /// <summary>
    ///     Initializes a new instance of the <see cref="GenericPropertyDefinition&lt;T&gt;" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="version">The version.</param>
    internal GenericPropertyDefinition(string xmlElementName, string uri, ExchangeVersion version)
        : base(xmlElementName, uri, version)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GenericPropertyDefinition&lt;T&gt;" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="flags">The flags.</param>
    /// <param name="version">The version.</param>
    internal GenericPropertyDefinition(
        string xmlElementName,
        string uri,
        PropertyDefinitionFlags flags,
        ExchangeVersion version
    )
        : base(xmlElementName, uri, flags, version)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GenericPropertyDefinition&lt;T&gt;" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="flags">The flags.</param>
    /// <param name="version">The version.</param>
    /// <param name="isNullable">if set to true, property value is nullable.</param>
    internal GenericPropertyDefinition(
        string xmlElementName,
        string uri,
        PropertyDefinitionFlags flags,
        ExchangeVersion version,
        bool isNullable
    )
        : base(xmlElementName, uri, flags, version, isNullable)
    {
    }

    /// <summary>
    ///     Parses the specified value.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>Value of string.</returns>
    internal override object Parse(string value)
    {
        return EwsUtilities.Parse<TPropertyValue>(value);
    }
}
