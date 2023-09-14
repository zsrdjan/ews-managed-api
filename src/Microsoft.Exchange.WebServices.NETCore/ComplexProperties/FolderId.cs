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
///     Represents the Id of a folder.
/// </summary>
[PublicAPI]
public sealed class FolderId : ServiceId
{
    private readonly WellKnownFolderName? _folderName;

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderId" /> class.
    /// </summary>
    internal FolderId()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderId" /> class. Use this constructor
    ///     to link this FolderId to an existing folder that you have the unique Id of.
    /// </summary>
    /// <param name="uniqueId">The unique Id used to initialize the FolderId.</param>
    public FolderId(string uniqueId)
        : base(uniqueId)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderId" /> class. Use this constructor
    ///     to link this FolderId to a well known folder (e.g. Inbox, Calendar or Contacts).
    /// </summary>
    /// <param name="folderName">The folder name used to initialize the FolderId.</param>
    public FolderId(WellKnownFolderName folderName)
    {
        _folderName = folderName;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderId" /> class. Use this constructor
    ///     to link this FolderId to a well known folder (e.g. Inbox, Calendar or Contacts) in a
    ///     specific mailbox.
    /// </summary>
    /// <param name="folderName">The folder name used to initialize the FolderId.</param>
    /// <param name="mailbox">The mailbox used to initialize the FolderId.</param>
    public FolderId(WellKnownFolderName folderName, Mailbox mailbox)
        : this(folderName)
    {
        Mailbox = mailbox;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return FolderName.HasValue ? XmlElementNames.DistinguishedFolderId : XmlElementNames.FolderId;
    }

    /// <summary>
    ///     Writes attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        if (FolderName.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Id, FolderName.Value.ToString().ToLowerInvariant());

            Mailbox?.WriteToXml(writer, XmlElementNames.Mailbox);
        }
        else
        {
            base.WriteAttributesToXml(writer);
        }
    }

    /// <summary>
    ///     Validates FolderId against a specified request version.
    /// </summary>
    /// <param name="version">The version.</param>
    internal void Validate(ExchangeVersion version)
    {
        // The FolderName property is a WellKnownFolderName, an enumeration type. If the property
        // is set, make sure that the value is valid for the request version.
        if (FolderName.HasValue)
        {
            EwsUtilities.ValidateEnumVersionValue(FolderName.Value, version);
        }
    }

    /// <summary>
    ///     Gets the name of the folder associated with the folder Id. Name and Id are mutually exclusive; if one is set, the
    ///     other is null.
    /// </summary>
    public WellKnownFolderName? FolderName => _folderName;

    /// <summary>
    ///     Gets the mailbox of the folder. Mailbox is only set when FolderName is set.
    /// </summary>
    public Mailbox? Mailbox { get; }

    /// <summary>
    ///     Defines an implicit conversion between string and FolderId.
    /// </summary>
    /// <param name="uniqueId">The unique Id to convert to FolderId.</param>
    /// <returns>A FolderId initialized with the specified unique Id.</returns>
    public static implicit operator FolderId(string uniqueId)
    {
        return new FolderId(uniqueId);
    }

    /// <summary>
    ///     Defines an implicit conversion between WellKnownFolderName and FolderId.
    /// </summary>
    /// <param name="folderName">The folder name to convert to FolderId.</param>
    /// <returns>A FolderId initialized with the specified folder name.</returns>
    public static implicit operator FolderId(WellKnownFolderName folderName)
    {
        return new FolderId(folderName);
    }

    /// <summary>
    ///     True if this instance is valid, false otherthise.
    /// </summary>
    /// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
    internal override bool IsValid
    {
        get
        {
            if (FolderName.HasValue)
            {
                return Mailbox == null || Mailbox.IsValid;
            }

            return base.IsValid;
        }
    }

    /// <summary>
    ///     Determines whether the specified <see cref="T:System.Object" /> is equal to the current
    ///     <see cref="T:System.Object" />.
    /// </summary>
    /// <param name="obj">The <see cref="T:System.Object" /> to compare with the current <see cref="T:System.Object" />.</param>
    /// <returns>
    ///     true if the specified <see cref="T:System.Object" /> is equal to the current <see cref="T:System.Object" />;
    ///     otherwise, false.
    /// </returns>
    /// <exception cref="T:System.NullReferenceException">The <paramref name="obj" /> parameter is null.</exception>
    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(this, obj))
        {
            return true;
        }

        if (obj is not FolderId other)
        {
            return false;
        }

        if (FolderName.HasValue)
        {
            if (other.FolderName.HasValue && FolderName.Value.Equals(other.FolderName.Value))
            {
                if (Mailbox != null)
                {
                    return Mailbox.Equals(other.Mailbox);
                }

                if (other.Mailbox == null)
                {
                    return true;
                }
            }
        }
        else if (base.Equals(other))
        {
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Serves as a hash function for a particular type.
    /// </summary>
    /// <returns>
    ///     A hash code for the current <see cref="T:System.Object" />.
    /// </returns>
    public override int GetHashCode()
    {
        int hashCode;

        if (FolderName.HasValue)
        {
            hashCode = FolderName.Value.GetHashCode();

            if (Mailbox != null && Mailbox.IsValid)
            {
                hashCode ^= Mailbox.GetHashCode();
            }
        }
        else
        {
            hashCode = base.GetHashCode();
        }

        return hashCode;
    }

    /// <summary>
    ///     Returns a <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </summary>
    /// <returns>
    ///     A <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </returns>
    public override string ToString()
    {
        if (IsValid)
        {
            if (FolderName.HasValue)
            {
                if (Mailbox != null && Mailbox.IsValid)
                {
                    return $"{_folderName.Value} ({Mailbox})";
                }

                return FolderName.Value.ToString();
            }

            return base.ToString();
        }

        return string.Empty;
    }
}
