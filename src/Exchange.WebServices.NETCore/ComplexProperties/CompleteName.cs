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
///     Represents the complete name of a contact.
/// </summary>
[PublicAPI]
public sealed class CompleteName : ComplexProperty
{
    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Title:
            {
                Title = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.FirstName:
            {
                GivenName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.MiddleName:
            {
                MiddleName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.LastName:
            {
                Surname = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.Suffix:
            {
                Suffix = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.Initials:
            {
                Initials = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.FullName:
            {
                FullName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.NickName:
            {
                NickName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.YomiFirstName:
            {
                YomiGivenName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.YomiLastName:
            {
                YomiSurname = reader.ReadElementValue();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes the elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Title, Title);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FirstName, GivenName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MiddleName, MiddleName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LastName, Surname);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Suffix, Suffix);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Initials, Initials);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FullName, FullName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.NickName, NickName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.YomiFirstName, YomiGivenName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.YomiLastName, YomiSurname);
    }


    #region Properties

    /// <summary>
    ///     Gets the contact's title.
    /// </summary>
    public string Title { get; private set; }

    /// <summary>
    ///     Gets the given name (first name) of the contact.
    /// </summary>
    public string GivenName { get; private set; }

    /// <summary>
    ///     Gets the middle name of the contact.
    /// </summary>
    public string MiddleName { get; private set; }

    /// <summary>
    ///     Gets the surname (last name) of the contact.
    /// </summary>
    public string Surname { get; private set; }

    /// <summary>
    ///     Gets the suffix of the contact.
    /// </summary>
    public string Suffix { get; private set; }

    /// <summary>
    ///     Gets the initials of the contact.
    /// </summary>
    public string Initials { get; private set; }

    /// <summary>
    ///     Gets the full name of the contact.
    /// </summary>
    public string FullName { get; private set; }

    /// <summary>
    ///     Gets the nickname of the contact.
    /// </summary>
    public string NickName { get; private set; }

    /// <summary>
    ///     Gets the Yomi given name (first name) of the contact.
    /// </summary>
    public string YomiGivenName { get; private set; }

    /// <summary>
    ///     Gets the Yomi surname (last name) of the contact.
    /// </summary>
    public string YomiSurname { get; private set; }

    #endregion
}
