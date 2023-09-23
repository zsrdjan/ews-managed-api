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

using System.Diagnostics.CodeAnalysis;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a simple property bag.
/// </summary>
/// <typeparam name="TKey">The type of the key.</typeparam>
internal class SimplePropertyBag<TKey> : IEnumerable<KeyValuePair<TKey, object>>
    where TKey : notnull
{
    private readonly List<TKey> _addedItems = new();
    private readonly Dictionary<TKey, object> _items = new();
    private readonly List<TKey> _modifiedItems = new();
    private readonly List<TKey> _removedItems = new();

    /// <summary>
    ///     Gets the added items.
    /// </summary>
    /// <value>The added items.</value>
    internal IEnumerable<TKey> AddedItems => _addedItems;

    /// <summary>
    ///     Gets the removed items.
    /// </summary>
    /// <value>The removed items.</value>
    internal IEnumerable<TKey> RemovedItems => _removedItems;

    /// <summary>
    ///     Gets the modified items.
    /// </summary>
    /// <value>The modified items.</value>
    internal IEnumerable<TKey> ModifiedItems => _modifiedItems;

    /// <summary>
    ///     Gets or sets the <see cref="object" /> with the specified key.
    /// </summary>
    /// <param name="key">Key.</param>
    /// <value>Value associated with key.</value>
    public object? this[TKey key]
    {
        get => TryGetValue(key, out var value) ? value : null;

        set
        {
            if (value == null)
            {
                InternalRemoveItem(key);
            }
            else
            {
                // If the item was to be deleted, the deletion becomes an update.
                if (_removedItems.Remove(key))
                {
                    InternalAddItemToChangeList(key, _modifiedItems);
                }
                else
                {
                    // If the property value was not set, we have a newly set property.
                    if (!ContainsKey(key))
                    {
                        InternalAddItemToChangeList(key, _addedItems);
                    }
                    else
                    {
                        // The last case is that we have a modified property.
                        if (!_modifiedItems.Contains(key))
                        {
                            InternalAddItemToChangeList(key, _modifiedItems);
                        }
                    }
                }

                _items[key] = value;
                Changed();
            }
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SimplePropertyBag&lt;TKey&gt;" /> class.
    /// </summary>
    public SimplePropertyBag()
    {
    }


    /// <summary>
    ///     Gets an enumerator that iterates through the elements of the collection.
    /// </summary>
    /// <returns>An IEnumerator for the collection.</returns>
    public IEnumerator<KeyValuePair<TKey, object>> GetEnumerator()
    {
        return _items.GetEnumerator();
    }

    /// <summary>
    ///     Gets an enumerator that iterates through the elements of the collection.
    /// </summary>
    /// <returns>An IEnumerator for the collection.</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return _items.GetEnumerator();
    }


    /// <summary>
    ///     Add item to change list.
    /// </summary>
    /// <param name="key">The key.</param>
    /// <param name="changeList">The change list.</param>
    private static void InternalAddItemToChangeList(TKey key, ICollection<TKey> changeList)
    {
        if (!changeList.Contains(key))
        {
            changeList.Add(key);
        }
    }

    /// <summary>
    ///     Triggers dispatch of the change event.
    /// </summary>
    private void Changed()
    {
        OnChange?.Invoke();
    }

    /// <summary>
    ///     Remove item.
    /// </summary>
    /// <param name="key">The key.</param>
    private void InternalRemoveItem(TKey key)
    {
        if (TryGetValue(key, out _))
        {
            _items.Remove(key);
            _removedItems.Add(key);
            Changed();
        }
    }

    /// <summary>
    ///     Clears the change log.
    /// </summary>
    public void ClearChangeLog()
    {
        _removedItems.Clear();
        _addedItems.Clear();
        _modifiedItems.Clear();
    }

    /// <summary>
    ///     Determines whether the specified key is in the property bag.
    /// </summary>
    /// <param name="key">The key.</param>
    /// <returns>
    ///     <c>true</c> if the specified key exists; otherwise, <c>false</c>.
    /// </returns>
    public bool ContainsKey(TKey key)
    {
        return _items.ContainsKey(key);
    }

    /// <summary>
    ///     Tries to get value.
    /// </summary>
    /// <param name="key">The key.</param>
    /// <param name="value">The value.</param>
    /// <returns>True if value exists in property bag.</returns>
    public bool TryGetValue(TKey key, [MaybeNullWhen(false)] out object value)
    {
        return _items.TryGetValue(key, out value);
    }

    /// <summary>
    ///     Occurs when Changed.
    /// </summary>
    public event PropertyBagChangedDelegate? OnChange;
}
