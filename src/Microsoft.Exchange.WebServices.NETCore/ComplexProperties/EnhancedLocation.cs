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
///     Represents Enhanced Location.
/// </summary>
[PublicAPI]
public sealed class EnhancedLocation : ComplexProperty
{
    private string _displayName;
    private string _annotation;
    private PersonaPostalAddress? _personaPostalAddress;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EnhancedLocation" /> class.
    /// </summary>
    internal EnhancedLocation()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EnhancedLocation" /> class.
    /// </summary>
    /// <param name="displayName">The location DisplayName.</param>
    public EnhancedLocation(string displayName)
        : this(displayName, string.Empty, new PersonaPostalAddress())
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EnhancedLocation" /> class.
    /// </summary>
    /// <param name="displayName">The location DisplayName.</param>
    /// <param name="annotation">The annotation on the location.</param>
    public EnhancedLocation(string displayName, string annotation)
        : this(displayName, annotation, new PersonaPostalAddress())
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EnhancedLocation" /> class.
    /// </summary>
    /// <param name="displayName">The location DisplayName.</param>
    /// <param name="annotation">The annotation on the location.</param>
    /// <param name="personaPostalAddress">The persona postal address.</param>
    public EnhancedLocation(string displayName, string annotation, PersonaPostalAddress personaPostalAddress)
        : this()
    {
        _displayName = displayName;
        _annotation = annotation;
        _personaPostalAddress = personaPostalAddress;
        _personaPostalAddress.OnChange += PersonaPostalAddress_OnChange;
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
            case XmlElementNames.LocationDisplayName:
            {
                _displayName = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.LocationAnnotation:
            {
                _annotation = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.PersonaPostalAddress:
            {
                _personaPostalAddress = new PersonaPostalAddress();
                _personaPostalAddress.LoadFromXml(reader);
                _personaPostalAddress.OnChange += PersonaPostalAddress_OnChange;
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Gets or sets the Location DisplayName.
    /// </summary>
    public string DisplayName
    {
        get => _displayName;
        set => SetFieldValue(ref _displayName, value);
    }

    /// <summary>
    ///     Gets or sets the Location Annotation.
    /// </summary>
    public string Annotation
    {
        get => _annotation;
        set => SetFieldValue(ref _annotation, value);
    }

    /// <summary>
    ///     Gets or sets the Persona Postal Address.
    /// </summary>
    public PersonaPostalAddress? PersonaPostalAddress
    {
        get => _personaPostalAddress;
        set
        {
            if (!_personaPostalAddress.Equals(value))
            {
                if (_personaPostalAddress != null)
                {
                    _personaPostalAddress.OnChange -= PersonaPostalAddress_OnChange;
                }

                SetFieldValue(ref _personaPostalAddress, value);

                _personaPostalAddress.OnChange += PersonaPostalAddress_OnChange;
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationDisplayName, _displayName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationAnnotation, _annotation);
        _personaPostalAddress.WriteToXml(writer);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void InternalValidate()
    {
        base.InternalValidate();
        EwsUtilities.ValidateParam(_displayName, "DisplayName");
        EwsUtilities.ValidateParamAllowNull(_annotation, "Annotation");
        EwsUtilities.ValidateParamAllowNull(_personaPostalAddress, "PersonaPostalAddress");
    }

    /// <summary>
    ///     PersonaPostalAddress OnChange.
    /// </summary>
    /// <param name="complexProperty">ComplexProperty object.</param>
    private void PersonaPostalAddress_OnChange(ComplexProperty complexProperty)
    {
        Changed();
    }
}
