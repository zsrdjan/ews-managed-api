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
using System.Reflection;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a collection of properties that can be sent to and retrieved from EWS.
/// </summary>
/// <typeparam name="TComplexProperty">ComplexProperty type.</typeparam>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class ComplexPropertyCollection<TComplexProperty> : ComplexProperty, IEnumerable<TComplexProperty>,
    ICustomUpdateSerializer
    where TComplexProperty : ComplexProperty
{
    /// <summary>
    ///     Creates the complex property.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Complex property instance.</returns>
    internal abstract TComplexProperty? CreateComplexProperty(string xmlElementName);

    /// <summary>
    ///     Gets the name of the collection item XML element.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    /// <returns>XML element name.</returns>
    internal abstract string GetCollectionItemXmlElementName(TComplexProperty complexProperty);

    /// <summary>
    ///     Initializes a new instance of the <see cref="ComplexPropertyCollection&lt;TComplexProperty&gt;" /> class.
    /// </summary>
    internal ComplexPropertyCollection()
    {
    }

    /// <summary>
    ///     Item changed.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    internal void ItemChanged(ComplexProperty complexProperty)
    {
        var property = complexProperty as TComplexProperty;

        EwsUtilities.Assert(
            property != null,
            "ComplexPropertyCollection.ItemChanged",
            $"ComplexPropertyCollection.ItemChanged: the type of the complexProperty argument ({complexProperty.GetType().Name}) is not supported."
        );

        if (!AddedItems.Contains(property))
        {
            if (!ModifiedItems.Contains(property))
            {
                ModifiedItems.Add(property);
                Changed();
            }
        }
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="localElementName">Name of the local element.</param>
    internal override void LoadFromXml(EwsServiceXmlReader reader, string localElementName)
    {
        LoadFromXml(reader, XmlNamespace.Types, localElementName);
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="localElementName">Name of the local element.</param>
    internal override void LoadFromXml(EwsServiceXmlReader reader, XmlNamespace xmlNamespace, string localElementName)
    {
        reader.EnsureCurrentNodeIsStartElement(xmlNamespace, localElementName);

        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    var complexProperty = CreateComplexProperty(reader.LocalName);

                    if (complexProperty != null)
                    {
                        complexProperty.LoadFromXml(reader, reader.LocalName);
                        InternalAdd(complexProperty, true);
                    }
                    else
                    {
                        reader.SkipCurrentElement();
                    }
                }
            } while (!reader.IsEndElement(xmlNamespace, localElementName));
        }
    }

    /// <summary>
    ///     Loads from XML to update itself.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal override void UpdateFromXml(EwsServiceXmlReader reader, XmlNamespace xmlNamespace, string xmlElementName)
    {
        reader.EnsureCurrentNodeIsStartElement(xmlNamespace, xmlElementName);

        if (!reader.IsEmptyElement)
        {
            var index = 0;
            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    var complexProperty = CreateComplexProperty(reader.LocalName);
                    var actualComplexProperty = this[index++];

                    if (complexProperty == null ||
                        !complexProperty.GetType()
                            .GetTypeInfo()
                            .IsAssignableFrom(actualComplexProperty.GetType().GetTypeInfo()))
                    {
                        throw new ServiceLocalException(Strings.PropertyTypeIncompatibleWhenUpdatingCollection);
                    }

                    actualComplexProperty.UpdateFromXml(reader, xmlNamespace, reader.LocalName);
                }
            } while (!reader.IsEndElement(xmlNamespace, xmlElementName));
        }
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal override void WriteToXml(EwsServiceXmlWriter writer, XmlNamespace xmlNamespace, string xmlElementName)
    {
        if (ShouldWriteToRequest())
        {
            base.WriteToXml(writer, xmlNamespace, xmlElementName);
        }
    }

    /// <summary>
    ///     Determine whether we should write collection to XML or not.
    /// </summary>
    /// <returns>True if collection contains at least one element.</returns>
    internal virtual bool ShouldWriteToRequest()
    {
        // Only write collection if it has at least one element.
        return Count > 0;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        foreach (var complexProperty in this)
        {
            complexProperty.WriteToXml(writer, GetCollectionItemXmlElementName(complexProperty));
        }
    }

    /// <summary>
    ///     Clears the change log.
    /// </summary>
    internal override void ClearChangeLog()
    {
        RemovedItems.Clear();
        AddedItems.Clear();
        ModifiedItems.Clear();
    }

    /// <summary>
    ///     Removes from change log.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    internal void RemoveFromChangeLog(TComplexProperty complexProperty)
    {
        RemovedItems.Remove(complexProperty);
        ModifiedItems.Remove(complexProperty);
        AddedItems.Remove(complexProperty);
    }

    /// <summary>
    ///     Gets the items.
    /// </summary>
    /// <value>The items.</value>
    internal List<TComplexProperty> Items { get; } = new();

    /// <summary>
    ///     Gets the added items.
    /// </summary>
    /// <value>The added items.</value>
    internal List<TComplexProperty> AddedItems { get; } = new();

    /// <summary>
    ///     Gets the modified items.
    /// </summary>
    /// <value>The modified items.</value>
    internal List<TComplexProperty> ModifiedItems { get; } = new();

    /// <summary>
    ///     Gets the removed items.
    /// </summary>
    /// <value>The removed items.</value>
    internal List<TComplexProperty> RemovedItems { get; } = new();

    /// <summary>
    ///     Add complex property.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    internal void InternalAdd(TComplexProperty complexProperty)
    {
        InternalAdd(complexProperty, false);
    }

    /// <summary>
    ///     Add complex property.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    /// <param name="loading">If true, collection is being loaded.</param>
    private void InternalAdd(TComplexProperty complexProperty, bool loading)
    {
        EwsUtilities.Assert(
            complexProperty != null,
            "ComplexPropertyCollection.InternalAdd",
            "complexProperty is null"
        );

        if (!Items.Contains(complexProperty))
        {
            Items.Add(complexProperty);
            if (!loading)
            {
                RemovedItems.Remove(complexProperty);
                AddedItems.Add(complexProperty);
            }

            complexProperty.OnChange += ItemChanged;
            Changed();
        }
    }

    /// <summary>
    ///     Clear collection.
    /// </summary>
    internal void InternalClear()
    {
        while (Count > 0)
        {
            InternalRemoveAt(0);
        }
    }

    /// <summary>
    ///     Remote entry at index.
    /// </summary>
    /// <param name="index">The index.</param>
    internal void InternalRemoveAt(int index)
    {
        EwsUtilities.Assert(
            index >= 0 && index < Count,
            "ComplexPropertyCollection.InternalRemoveAt",
            "index is out of range."
        );

        InternalRemove(Items[index]);
    }

    /// <summary>
    ///     Remove specified complex property.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    /// <returns>True if the complex property was successfully removed from the collection, false otherwise.</returns>
    internal bool InternalRemove(TComplexProperty complexProperty)
    {
        EwsUtilities.Assert(
            complexProperty != null,
            "ComplexPropertyCollection.InternalRemove",
            "complexProperty is null"
        );

        if (Items.Remove(complexProperty))
        {
            complexProperty.OnChange -= ItemChanged;

            if (!AddedItems.Contains(complexProperty))
            {
                RemovedItems.Add(complexProperty);
            }
            else
            {
                AddedItems.Remove(complexProperty);
            }

            ModifiedItems.Remove(complexProperty);
            Changed();
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Determines whether a specific property is in the collection.
    /// </summary>
    /// <param name="complexProperty">The property to locate in the collection.</param>
    /// <returns>True if the property was found in the collection, false otherwise.</returns>
    public bool Contains(TComplexProperty complexProperty)
    {
        return Items.Contains(complexProperty);
    }

    /// <summary>
    ///     Searches for a specific property and return its zero-based index within the collection.
    /// </summary>
    /// <param name="complexProperty">The property to locate in the collection.</param>
    /// <returns>The zero-based index of the property within the collection.</returns>
    public int IndexOf(TComplexProperty complexProperty)
    {
        return Items.IndexOf(complexProperty);
    }

    /// <summary>
    ///     Gets the total number of properties in the collection.
    /// </summary>
    public int Count => Items.Count;

    /// <summary>
    ///     Gets the property at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the property to get.</param>
    /// <returns>The property at the specified index.</returns>
    public TComplexProperty this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
            }

            return Items[index];
        }
    }


    #region IEnumerable<TComplexProperty> Members

    /// <summary>
    ///     Gets an enumerator that iterates through the elements of the collection.
    /// </summary>
    /// <returns>An IEnumerator for the collection.</returns>
    public IEnumerator<TComplexProperty> GetEnumerator()
    {
        return Items.GetEnumerator();
    }

    #endregion


    #region IEnumerable Members

    /// <summary>
    ///     Gets an enumerator that iterates through the elements of the collection.
    /// </summary>
    /// <returns>An IEnumerator for the collection.</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return Items.GetEnumerator();
    }

    #endregion


    #region ICustomXmlUpdateSerializer Members

    /// <summary>
    ///     Writes the update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <param name="propertyDefinition">Property definition.</param>
    /// <returns>True if property generated serialization.</returns>
    bool ICustomUpdateSerializer.WriteSetUpdateToXml(
        EwsServiceXmlWriter writer,
        ServiceObject ewsObject,
        PropertyDefinition propertyDefinition
    )
    {
        // If the collection is empty, delete the property.
        if (Count == 0)
        {
            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
            propertyDefinition.WriteToXml(writer);
            writer.WriteEndElement();
            return true;
        }

        // Otherwise, use the default XML serializer.

        return false;
    }

    /// <summary>
    ///     Writes the deletion update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <returns>True if property generated serialization.</returns>
    bool ICustomUpdateSerializer.WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
    {
        // Use the default XML serializer.
        return false;
    }

    #endregion
}
