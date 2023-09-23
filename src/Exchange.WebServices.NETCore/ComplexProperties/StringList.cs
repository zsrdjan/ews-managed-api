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
///     Represents a list of strings.
/// </summary>
[PublicAPI]
public sealed class StringList : ComplexProperty, IEnumerable<string>
{
    private readonly List<string> _items = new();

    private readonly string _itemXmlElementName = XmlElementNames.String;

    /// <summary>
    ///     Gets the number of strings in the list.
    /// </summary>
    public int Count => _items.Count;

    /// <summary>
    ///     Gets or sets the string at the specified index.
    /// </summary>
    /// <param name="index">The index of the string to get or set.</param>
    /// <returns>The string at the specified index.</returns>
    public string this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
            }

            return _items[index];
        }

        set
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
            }

            if (_items[index] != value)
            {
                _items[index] = value;
                Changed();
            }
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="StringList" /> class.
    /// </summary>
    public StringList()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="StringList" /> class.
    /// </summary>
    /// <param name="strings">The strings.</param>
    public StringList(IEnumerable<string> strings)
    {
        AddRange(strings);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="StringList" /> class.
    /// </summary>
    /// <param name="itemXmlElementName">Name of the item XML element.</param>
    internal StringList(string itemXmlElementName)
    {
        _itemXmlElementName = itemXmlElementName;
    }


    #region IEnumerable<string> Members

    /// <summary>
    ///     Gets an enumerator that iterates through the elements of the collection.
    /// </summary>
    /// <returns>An IEnumerator for the collection.</returns>
    public IEnumerator<string> GetEnumerator()
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
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        if (reader.LocalName == _itemXmlElementName)
        {
            Add(reader.ReadValue());

            return true;
        }

        return false;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        foreach (var item in this)
        {
            writer.WriteStartElement(XmlNamespace.Types, _itemXmlElementName);
            writer.WriteValue(item, _itemXmlElementName);
            writer.WriteEndElement();
        }
    }

    /// <summary>
    ///     Adds a string to the list.
    /// </summary>
    /// <param name="s">The string to add.</param>
    public void Add(string s)
    {
        _items.Add(s);
        Changed();
    }

    /// <summary>
    ///     Adds multiple strings to the list.
    /// </summary>
    /// <param name="strings">The strings to add.</param>
    public void AddRange(IEnumerable<string> strings)
    {
        var changed = false;

        foreach (var s in strings)
        {
            if (!Contains(s))
            {
                _items.Add(s);
                changed = true;
            }
        }

        if (changed)
        {
            Changed();
        }
    }

    /// <summary>
    ///     Determines whether the list contains a specific string.
    /// </summary>
    /// <param name="s">The string to check the presence of.</param>
    /// <returns>True if s is present in the list, false otherwise.</returns>
    public bool Contains(string s)
    {
        return _items.Contains(s);
    }

    /// <summary>
    ///     Removes a string from the list.
    /// </summary>
    /// <param name="s">The string to remove.</param>
    /// <returns>True is s was removed, false otherwise.</returns>
    public bool Remove(string s)
    {
        var result = _items.Remove(s);

        if (result)
        {
            Changed();
        }

        return result;
    }

    /// <summary>
    ///     Removes the string at the specified position from the list.
    /// </summary>
    /// <param name="index">The index of the string to remove.</param>
    public void RemoveAt(int index)
    {
        if (index < 0 || index >= Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
        }

        _items.RemoveAt(index);

        Changed();
    }

    /// <summary>
    ///     Clears the list.
    /// </summary>
    public void Clear()
    {
        _items.Clear();
        Changed();
    }

    /// <summary>
    ///     Generates a string representation of all the items in the list.
    /// </summary>
    /// <returns>A comma-separated list of the strings present in the list.</returns>
    public override string ToString()
    {
        return string.Join(",", _items.ToArray());
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
        if (obj is StringList other)
        {
            return ToString().Equals(other.ToString());
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
        return ToString().GetHashCode();
    }
}
