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
///     Represents a response to an individual item retrieval operation.
/// </summary>
public sealed class GetItemResponse : ServiceResponse
{
    private Item? item;
    private readonly PropertySet propertySet;

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetItemResponse" /> class.
    /// </summary>
    /// <param name="item">The item.</param>
    /// <param name="propertySet">The property set.</param>
    internal GetItemResponse(Item? item, PropertySet propertySet)
    {
        this.item = item;
        this.propertySet = propertySet;

        EwsUtilities.Assert(this.propertySet != null, "GetItemResponse.ctor", "PropertySet should not be null");
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        base.ReadElementsFromXml(reader);

        var items = reader.ReadServiceObjectsCollectionFromXml(
            XmlElementNames.Items,
            GetObjectInstance,
            true, /* clearPropertyBag */
            propertySet, /* requestedPropertySet */
            false
        ); /* summaryPropertiesOnly */

        item = items[0];
    }

    /// <summary>
    ///     Gets Item instance.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Item.</returns>
    private Item? GetObjectInstance(ExchangeService service, string xmlElementName)
    {
        if (Item != null)
        {
            return Item;
        }

        return EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(service, xmlElementName);
    }

    /// <summary>
    ///     Gets the item that was retrieved.
    /// </summary>
    public Item? Item => item;
}
