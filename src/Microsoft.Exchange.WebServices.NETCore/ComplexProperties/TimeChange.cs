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
///     Represents a change of time for a time zone.
/// </summary>
internal sealed class TimeChange : ComplexProperty
{
    private string _timeZoneName;
    private TimeSpan? _offset;
    private Time? _time;
    private DateTime? _absoluteDate;
    private TimeChangeRecurrence? _recurrence;

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeChange" /> class.
    /// </summary>
    public TimeChange()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeChange" /> class.
    /// </summary>
    /// <param name="offset">The offset since the beginning of the year when the change occurs.</param>
    public TimeChange(TimeSpan offset)
        : this()
    {
        _offset = offset;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeChange" /> class.
    /// </summary>
    /// <param name="offset">The offset since the beginning of the year when the change occurs.</param>
    /// <param name="time">The time at which the change occurs.</param>
    public TimeChange(TimeSpan offset, Time time)
        : this(offset)
    {
        _time = time;
    }

    /// <summary>
    ///     Gets or sets the name of the associated time zone.
    /// </summary>
    public string TimeZoneName
    {
        get => _timeZoneName;
        set => SetFieldValue(ref _timeZoneName, value);
    }

    /// <summary>
    ///     Gets or sets the offset since the beginning of the year when the change occurs.
    /// </summary>
    public TimeSpan? Offset
    {
        get => _offset;
        set => SetFieldValue(ref _offset, value);
    }

    /// <summary>
    ///     Gets or sets the time at which the change occurs.
    /// </summary>
    public Time? Time
    {
        get => _time;
        set => SetFieldValue(ref _time, value);
    }

    /// <summary>
    ///     Gets or sets the absolute date at which the change occurs. AbsoluteDate and Recurrence are mutually exclusive;
    ///     setting one resets the other.
    /// </summary>
    public DateTime? AbsoluteDate
    {
        get => _absoluteDate;

        set
        {
            SetFieldValue(ref _absoluteDate, value);

            if (_absoluteDate.HasValue)
            {
                _recurrence = null;
            }
        }
    }

    /// <summary>
    ///     Gets or sets the recurrence pattern defining when the change occurs. Recurrence and AbsoluteDate are mutually
    ///     exclusive; setting one resets the other.
    /// </summary>
    public TimeChangeRecurrence? Recurrence
    {
        get => _recurrence;

        set
        {
            SetFieldValue(ref _recurrence, value);

            if (_recurrence != null)
            {
                _absoluteDate = null;
            }
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
            case XmlElementNames.Offset:
            {
                _offset = EwsUtilities.XsDurationToTimeSpan(reader.ReadElementValue());
                return true;
            }
            case XmlElementNames.RelativeYearlyRecurrence:
            {
                Recurrence = new TimeChangeRecurrence();
                Recurrence.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.AbsoluteDate:
            {
                var dateTime = DateTime.Parse(reader.ReadElementValue());

                // TODO: BUG
                _absoluteDate = new DateTime(dateTime.ToUniversalTime().Ticks, DateTimeKind.Unspecified);
                return true;
            }
            case XmlElementNames.Time:
            {
                _time = new Time(DateTime.Parse(reader.ReadElementValue()));
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
        _timeZoneName = reader.ReadAttributeValue(XmlAttributeNames.TimeZoneName);
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.TimeZoneName, TimeZoneName);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (Offset.HasValue)
        {
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Offset,
                EwsUtilities.TimeSpanToXsDuration(Offset.Value)
            );
        }

        if (Recurrence != null)
        {
            Recurrence.WriteToXml(writer, XmlElementNames.RelativeYearlyRecurrence);
        }

        if (AbsoluteDate.HasValue)
        {
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.AbsoluteDate,
                EwsUtilities.DateTimeToXsDate(new DateTime(AbsoluteDate.Value.Ticks, DateTimeKind.Unspecified))
            );
        }

        if (Time != null)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Time, Time.ToXsTime());
        }
    }
}
