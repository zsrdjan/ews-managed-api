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
///     Represents a recurrence pattern for a time change in a time zone.
/// </summary>
internal sealed class TimeChangeRecurrence : ComplexProperty
{
    private DayOfTheWeek? _dayOfTheWeek;
    private DayOfTheWeekIndex? _dayOfTheWeekIndex;
    private Month? _month;

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeChangeRecurrence" /> class.
    /// </summary>
    public TimeChangeRecurrence()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeChangeRecurrence" /> class.
    /// </summary>
    /// <param name="dayOfTheWeekIndex">The index of the day in the month at which the time change occurs.</param>
    /// <param name="dayOfTheWeek">The day of the week the time change occurs.</param>
    /// <param name="month">The month the time change occurs.</param>
    public TimeChangeRecurrence(DayOfTheWeekIndex dayOfTheWeekIndex, DayOfTheWeek dayOfTheWeek, Month month)
        : this()
    {
        _dayOfTheWeekIndex = dayOfTheWeekIndex;
        _dayOfTheWeek = dayOfTheWeek;
        _month = month;
    }

    /// <summary>
    ///     Gets or sets the index of the day in the month at which the time change occurs.
    /// </summary>
    public DayOfTheWeekIndex? DayOfTheWeekIndex
    {
        get => _dayOfTheWeekIndex;
        set => SetFieldValue(ref _dayOfTheWeekIndex, value);
    }

    /// <summary>
    ///     Gets or sets the day of the week the time change occurs.
    /// </summary>
    public DayOfTheWeek? DayOfTheWeek
    {
        get => _dayOfTheWeek;
        set => SetFieldValue(ref _dayOfTheWeek, value);
    }

    /// <summary>
    ///     Gets or sets the month the time change occurs.
    /// </summary>
    public Month? Month
    {
        get => _month;
        set => SetFieldValue(ref _month, value);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (DayOfTheWeek.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DaysOfWeek, DayOfTheWeek.Value);
        }

        if (_dayOfTheWeekIndex.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DayOfWeekIndex, DayOfTheWeekIndex.Value);
        }

        if (Month.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Month, Month.Value);
        }
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
            case XmlElementNames.DaysOfWeek:
            {
                _dayOfTheWeek = reader.ReadElementValue<DayOfTheWeek>();
                return true;
            }
            case XmlElementNames.DayOfWeekIndex:
            {
                _dayOfTheWeekIndex = reader.ReadElementValue<DayOfTheWeekIndex>();
                return true;
            }
            case XmlElementNames.Month:
            {
                _month = reader.ReadElementValue<Month>();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
