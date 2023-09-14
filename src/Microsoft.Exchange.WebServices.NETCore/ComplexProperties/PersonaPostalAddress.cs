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

using System.Xml;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents PersonaPostalAddress.
/// </summary>
[PublicAPI]
public sealed class PersonaPostalAddress : ComplexProperty
{
    private string _street;
    private string _city;
    private string _state;
    private string _country;
    private string _postalCode;
    private string _postOfficeBox;
    private string _type;
    private double? _latitude;
    private double? _longitude;
    private double? _accuracy;
    private double? _altitude;
    private double? _altitudeAccuracy;
    private string _formattedAddress;
    private string _uri;
    private LocationSource _source;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PersonaPostalAddress" /> class.
    /// </summary>
    internal PersonaPostalAddress()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PersonaPostalAddress" /> class.
    /// </summary>
    /// <param name="street">The Street Address.</param>
    /// <param name="city">The City value.</param>
    /// <param name="state">The State value.</param>
    /// <param name="country">The country value.</param>
    /// <param name="postalCode">The postal code value.</param>
    /// <param name="postOfficeBox">The Post Office Box.</param>
    /// <param name="locationSource">The location Source.</param>
    /// <param name="locationUri">The location Uri.</param>
    /// <param name="formattedAddress">The location street Address in formatted address.</param>
    /// <param name="latitude">The location latitude.</param>
    /// <param name="longitude">The location longitude.</param>
    /// <param name="accuracy">The location accuracy.</param>
    /// <param name="altitude">The location altitude.</param>
    /// <param name="altitudeAccuracy">The location altitude Accuracy.</param>
    public PersonaPostalAddress(
        string street,
        string city,
        string state,
        string country,
        string postalCode,
        string postOfficeBox,
        LocationSource locationSource,
        string locationUri,
        string formattedAddress,
        double latitude,
        double longitude,
        double accuracy,
        double altitude,
        double altitudeAccuracy
    )
        : this()
    {
        _street = street;
        _city = city;
        _state = state;
        _country = country;
        _postalCode = postalCode;
        _postOfficeBox = postOfficeBox;
        _latitude = latitude;
        _longitude = longitude;
        _source = locationSource;
        _uri = locationUri;
        _formattedAddress = formattedAddress;
        _accuracy = accuracy;
        _altitude = altitude;
        _altitudeAccuracy = altitudeAccuracy;
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
            case XmlElementNames.Street:
            {
                _street = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.City:
            {
                _city = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.State:
            {
                _state = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.Country:
            {
                _country = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.PostalCode:
            {
                _postalCode = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.PostOfficeBox:
            {
                _postOfficeBox = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.PostalAddressType:
            {
                _type = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.Latitude:
            {
                _latitude = reader.ReadValue<double>();
                return true;
            }
            case XmlElementNames.Longitude:
            {
                _longitude = reader.ReadValue<double>();
                return true;
            }
            case XmlElementNames.Accuracy:
            {
                _accuracy = reader.ReadValue<double>();
                return true;
            }
            case XmlElementNames.Altitude:
            {
                _altitude = reader.ReadValue<double>();
                return true;
            }
            case XmlElementNames.AltitudeAccuracy:
            {
                _altitudeAccuracy = reader.ReadValue<double>();
                return true;
            }
            case XmlElementNames.FormattedAddress:
            {
                _formattedAddress = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.LocationUri:
            {
                _uri = reader.ReadValue<string>();
                return true;
            }
            case XmlElementNames.LocationSource:
            {
                _source = reader.ReadValue<LocationSource>();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        do
        {
            reader.Read();

            if (reader.NodeType == XmlNodeType.Element)
            {
                TryReadElementFromXml(reader);
            }
        } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.PersonaPostalAddress));
    }

    /// <summary>
    ///     Gets or sets the street.
    /// </summary>
    public string Street
    {
        get => _street;
        set => SetFieldValue(ref _street, value);
    }

    /// <summary>
    ///     Gets or sets the City.
    /// </summary>
    public string City
    {
        get => _city;
        set => SetFieldValue(ref _city, value);
    }

    /// <summary>
    ///     Gets or sets the state.
    /// </summary>
    public string State
    {
        get => _state;
        set => SetFieldValue(ref _state, value);
    }

    /// <summary>
    ///     Gets or sets the Country.
    /// </summary>
    public string Country
    {
        get => _country;
        set => SetFieldValue(ref _country, value);
    }

    /// <summary>
    ///     Gets or sets the postalCode.
    /// </summary>
    public string PostalCode
    {
        get => _postalCode;
        set => SetFieldValue(ref _postalCode, value);
    }

    /// <summary>
    ///     Gets or sets the postOfficeBox.
    /// </summary>
    public string PostOfficeBox
    {
        get => _postOfficeBox;
        set => SetFieldValue(ref _postOfficeBox, value);
    }

    /// <summary>
    ///     Gets or sets the type.
    /// </summary>
    public string Type
    {
        get => _type;
        set => SetFieldValue(ref _type, value);
    }

    /// <summary>
    ///     Gets or sets the location source type.
    /// </summary>
    public LocationSource Source
    {
        get => _source;
        set => SetFieldValue(ref _source, value);
    }

    /// <summary>
    ///     Gets or sets the location Uri.
    /// </summary>
    public string Uri
    {
        get => _uri;
        set => SetFieldValue(ref _uri, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating location latitude.
    /// </summary>
    public double? Latitude
    {
        get => _latitude;
        set => SetFieldValue(ref _latitude, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating location longitude.
    /// </summary>
    public double? Longitude
    {
        get => _longitude;
        set => SetFieldValue(ref _longitude, value);
    }

    /// <summary>
    ///     Gets or sets the location accuracy.
    /// </summary>
    public double? Accuracy
    {
        get => _accuracy;
        set => SetFieldValue(ref _accuracy, value);
    }

    /// <summary>
    ///     Gets or sets the location altitude.
    /// </summary>
    public double? Altitude
    {
        get => _altitude;
        set => SetFieldValue(ref _altitude, value);
    }

    /// <summary>
    ///     Gets or sets the location altitude accuracy.
    /// </summary>
    public double? AltitudeAccuracy
    {
        get => _altitudeAccuracy;
        set => SetFieldValue(ref _altitudeAccuracy, value);
    }

    /// <summary>
    ///     Gets or sets the street address.
    /// </summary>
    public string FormattedAddress
    {
        get => _formattedAddress;
        set => SetFieldValue(ref _formattedAddress, value);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Street, _street);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.City, _city);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.State, _state);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Country, _country);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostalCode, _postalCode);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostOfficeBox, _postOfficeBox);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostalAddressType, _type);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Latitude, _latitude);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Longitude, _longitude);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Accuracy, _accuracy);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Altitude, _altitude);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AltitudeAccuracy, _altitudeAccuracy);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FormattedAddress, _formattedAddress);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationUri, _uri);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationSource, _source);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.PersonaPostalAddress);

        WriteElementsToXml(writer);

        writer.WriteEndElement(); // xmlElementName
    }
}
