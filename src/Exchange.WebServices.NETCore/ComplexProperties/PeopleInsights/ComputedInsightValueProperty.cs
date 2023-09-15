// ---------------------------------------------------------------------------
// <copyright file="ComputedInsightValueProperty.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements the class for computed insight value property.</summary>
//-----------------------------------------------------------------------

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a computed insight value.
/// </summary>
[PublicAPI]
public sealed class ComputedInsightValueProperty : ComplexProperty
{
    private string _key;
    private string _value;

    /// <summary>
    ///     Gets or sets the Key
    /// </summary>
    public string Key
    {
        get => _key;
        set => SetFieldValue(ref _key, value);
    }

    /// <summary>
    ///     Gets or sets the Value
    /// </summary>
    public string Value
    {
        get => _value;
        set => SetFieldValue(ref _value, value);
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">XML reader</param>
    /// <returns>Whether the element was read</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Key:
            {
                Key = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Value:
            {
                Value = reader.ReadElementValue();
                break;
            }
            default:
            {
                return false;
            }
        }

        return true;
    }
}
