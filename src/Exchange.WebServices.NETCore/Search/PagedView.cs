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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a view settings that support paging in a search operation.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class PagedView : ViewBase
{
    private int _offset;
    private int _pageSize;

    /// <summary>
    ///     The maximum number of items or folders the search operation should return.
    /// </summary>
    public int PageSize
    {
        get => _pageSize;

        set
        {
            if (value <= 0)
            {
                throw new ArgumentException(Strings.ValueMustBeGreaterThanZero);
            }

            _pageSize = value;
        }
    }

    /// <summary>
    ///     Gets or sets the base point of the offset.
    /// </summary>
    public OffsetBasePoint OffsetBasePoint { get; set; } = OffsetBasePoint.Beginning;

    /// <summary>
    ///     Gets or sets the offset.
    /// </summary>
    public int Offset
    {
        get => _offset;

        set
        {
            if (value < 0)
            {
                throw new ArgumentException(Strings.OffsetMustBeGreaterThanZero);
            }

            _offset = value;
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PagedView" /> class.
    /// </summary>
    /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
    internal PagedView(int pageSize)
    {
        PageSize = pageSize;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PagedView" /> class.
    /// </summary>
    /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
    /// <param name="offset">The offset of the view from the base point.</param>
    internal PagedView(int pageSize, int offset)
        : this(pageSize)
    {
        Offset = offset;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PagedView" /> class.
    /// </summary>
    /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
    /// <param name="offset">The offset of the view from the base point.</param>
    /// <param name="offsetBasePoint">The base point of the offset.</param>
    internal PagedView(int pageSize, int offset, OffsetBasePoint offsetBasePoint)
        : this(pageSize, offset)
    {
        OffsetBasePoint = offsetBasePoint;
    }

    /// <summary>
    ///     Write to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void InternalWriteViewToXml(EwsServiceXmlWriter writer)
    {
        base.InternalWriteViewToXml(writer);

        writer.WriteAttributeValue(XmlAttributeNames.Offset, Offset);
        writer.WriteAttributeValue(XmlAttributeNames.BasePoint, OffsetBasePoint);
    }

    /// <summary>
    ///     Gets the maximum number of items or folders the search operation should return.
    /// </summary>
    /// <returns>The maximum number of items or folders that should be returned by the search operation.</returns>
    internal override int? GetMaxEntriesReturned()
    {
        return PageSize;
    }

    /// <summary>
    ///     Internals the write search settings to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="groupBy">The group by clause.</param>
    internal override void InternalWriteSearchSettingsToXml(EwsServiceXmlWriter writer, Grouping groupBy)
    {
        if (groupBy != null)
        {
            groupBy.WriteToXml(writer);
        }
    }

    /// <summary>
    ///     Writes OrderBy property to XML.
    /// </summary>
    /// <param name="writer">The writer</param>
    internal override void WriteOrderByToXml(EwsServiceXmlWriter writer)
    {
        // No order by for paged view
    }
}
