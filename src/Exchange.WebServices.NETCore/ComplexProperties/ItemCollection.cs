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

using System.ComponentModel;
using System.Xml;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a collection of items.
/// </summary>
/// <typeparam name="TItem">The type of item the collection contains.</typeparam>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ItemCollection<TItem> : ComplexProperty, IEnumerable<TItem>
    where TItem : Item
{
    private readonly List<TItem> _items = new();

    /// <summary>
    ///     Gets the total number of items in the collection.
    /// </summary>
    public int Count => _items.Count;

    /// <summary>
    ///     Gets the item at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the item to get.</param>
    /// <returns>The item at the specified index.</returns>
    public TItem this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
            }

            return _items[index];
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ItemCollection&lt;TItem&gt;" /> class.
    /// </summary>
    internal ItemCollection()
    {
    }


    #region IEnumerable<TItem> Members

    /// <summary>
    ///     Gets an enumerator that iterates through the elements of the collection.
    /// </summary>
    /// <returns>An IEnumerator for the collection.</returns>
    public IEnumerator<TItem> GetEnumerator()
    {
        return _items.GetEnumerator();
    }

    #endregion


    #region IEnumerable Members

    /// <summary>
    ///     Gets an enumerator that iterates through the elements of the collection.
    /// </summary>
    /// <returns>An IEnumerator for the collection.</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return _items.GetEnumerator();
    }

    #endregion


    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="localElementName">Name of the local element.</param>
    internal override void LoadFromXml(EwsServiceXmlReader reader, string localElementName)
    {
        reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, localElementName);
        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(reader.Service, reader.LocalName) is not
                        TItem item)
                    {
                        reader.SkipCurrentElement();
                    }
                    else
                    {
                        item.LoadFromXml(reader, true /* clearPropertyBag */);

                        _items.Add(item);
                    }
                }
            } while (!reader.IsEndElement(XmlNamespace.Types, localElementName));
        }
    }
}
