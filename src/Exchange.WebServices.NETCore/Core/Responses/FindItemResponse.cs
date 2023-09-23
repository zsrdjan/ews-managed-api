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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the response to a item search operation.
/// </summary>
/// <typeparam name="TItem">The type of items that the operation returned.</typeparam>
internal sealed class FindItemResponse<TItem> : ServiceResponse
    where TItem : Item
{
    private readonly bool _isGrouped;
    private readonly PropertySet _propertySet;

    /// <summary>
    ///     Gets a grouped list of items matching the specified search criteria that were found in Exchange. ItemGroups is
    ///     null if the search operation did not specify grouping options.
    /// </summary>
    public GroupedFindItemsResults<TItem> GroupedFindResults { get; private set; }

    /// <summary>
    ///     Gets the results of the search operation.
    /// </summary>
    public FindItemsResults<TItem> Results { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FindItemResponse&lt;TItem&gt;" /> class.
    /// </summary>
    /// <param name="isGrouped">if set to <c>true</c> if grouped.</param>
    /// <param name="propertySet">The property set.</param>
    internal FindItemResponse(bool isGrouped, PropertySet propertySet)
    {
        _isGrouped = isGrouped;
        _propertySet = propertySet;

        EwsUtilities.Assert(_propertySet != null, "FindItemResponse.ctor", "PropertySet should not be null");
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.RootFolder);

        var totalItemsInView = reader.ReadAttributeValue<int>(XmlAttributeNames.TotalItemsInView);
        var moreItemsAvailable = !reader.ReadAttributeValue<bool>(XmlAttributeNames.IncludesLastItemInRange);

        // Ignore IndexedPagingOffset attribute if moreItemsAvailable is false.
        var nextPageOffset = moreItemsAvailable
            ? reader.ReadNullableAttributeValue<int>(XmlAttributeNames.IndexedPagingOffset) : null;

        if (!_isGrouped)
        {
            Results = new FindItemsResults<TItem>
            {
                TotalCount = totalItemsInView,
                NextPageOffset = nextPageOffset,
                MoreAvailable = moreItemsAvailable,
            };
            InternalReadItemsFromXml(reader, _propertySet, Results.Items);
        }
        else
        {
            GroupedFindResults = new GroupedFindItemsResults<TItem>
            {
                TotalCount = totalItemsInView,
                NextPageOffset = nextPageOffset,
                MoreAvailable = moreItemsAvailable,
            };

            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Groups);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.GroupedItems))
                    {
                        var groupIndex = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.GroupIndex);

                        var itemList = new List<TItem>();
                        InternalReadItemsFromXml(reader, _propertySet, itemList);

                        reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.GroupedItems);

                        GroupedFindResults.ItemGroups.Add(new ItemGroup<TItem>(groupIndex, itemList));
                    }
                } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Groups));
            }
        }

        reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.RootFolder);

        reader.Read();

        if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms) && !reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    var term = new HighlightTerm();

                    term.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.HighlightTerm);
                    Results.HighlightTerms.Add(term);
                }
            } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms));
        }
    }

    /// <summary>
    ///     Read items from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="propertySet">The property set.</param>
    /// <param name="destinationList">The list in which to add the read items.</param>
    private static void InternalReadItemsFromXml(
        EwsServiceXmlReader reader,
        PropertySet propertySet,
        IList<TItem> destinationList
    )
    {
        EwsUtilities.Assert(
            destinationList != null,
            "FindItemResponse.InternalReadItemsFromXml",
            "destinationList is null."
        );

        reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Items);
        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    var item = EwsUtilities.CreateEwsObjectFromXmlElementName<TItem>(reader.Service, reader.LocalName);

                    if (item == null)
                    {
                        reader.SkipCurrentElement();
                    }
                    else
                    {
                        item.LoadFromXml(reader, true, propertySet, true);

                        destinationList.Add(item);
                    }
                }
            } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Items));
        }
    }

    /// <summary>
    ///     Creates an item instance.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Item</returns>
    private TItem? CreateItemInstance(ExchangeService service, string xmlElementName)
    {
        return EwsUtilities.CreateEwsObjectFromXmlElementName<TItem>(service, xmlElementName);
    }
}
