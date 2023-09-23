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
///     Represents the parameters associated with a search folder.
/// </summary>
[PublicAPI]
public sealed class SearchFolderParameters : ComplexProperty
{
    private SearchFilter? _searchFilter;
    private SearchFolderTraversal _traversal;

    /// <summary>
    ///     Gets or sets the traversal mode for the search folder.
    /// </summary>
    public SearchFolderTraversal Traversal
    {
        get => _traversal;
        set => SetFieldValue(ref _traversal, value);
    }

    /// <summary>
    ///     Gets the list of root folders the search folder searches in.
    /// </summary>
    public FolderIdCollection RootFolderIds { get; } = new();

    /// <summary>
    ///     Gets or sets the search filter associated with the search folder. Available search filter classes include
    ///     SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.
    /// </summary>
    public SearchFilter? SearchFilter
    {
        get => _searchFilter;

        set
        {
            if (_searchFilter != null)
            {
                _searchFilter.OnChange -= PropertyChanged;
            }

            SetFieldValue(ref _searchFilter, value);

            if (_searchFilter != null)
            {
                _searchFilter.OnChange += PropertyChanged;
            }
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SearchFolderParameters" /> class.
    /// </summary>
    internal SearchFolderParameters()
    {
        RootFolderIds.OnChange += PropertyChanged;
    }

    /// <summary>
    ///     Property changed.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    private void PropertyChanged(ComplexProperty complexProperty)
    {
        Changed();
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
            case XmlElementNames.BaseFolderIds:
            {
                RootFolderIds.InternalClear();
                RootFolderIds.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.Restriction:
            {
                reader.Read();
                _searchFilter = SearchFilter.LoadFromXml(reader);
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Reads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        Traversal = reader.ReadAttributeValue<SearchFolderTraversal>(XmlAttributeNames.Traversal);
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.Traversal, Traversal);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (SearchFilter != null)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Restriction);
            SearchFilter.WriteToXml(writer);
            writer.WriteEndElement(); // Restriction
        }

        RootFolderIds.WriteToXml(writer, XmlElementNames.BaseFolderIds);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal void Validate()
    {
        // Search folder must have at least one root folder id.
        if (RootFolderIds.Count == 0)
        {
            throw new ServiceValidationException(Strings.SearchParametersRootFolderIdsEmpty);
        }

        // Validate the search filter
        SearchFilter?.InternalValidate();
    }
}
