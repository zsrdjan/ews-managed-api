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

using Microsoft.Exchange.WebServices.Data.Misc;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a custom time zone time change.
/// </summary>
internal sealed class LegacyAvailabilityTimeZoneTime : ComplexProperty
{
    private TimeSpan _delta;
    private TimeSpan _timeOfDay;

    /// <summary>
    ///     Gets if current time presents DST transition time
    /// </summary>
    internal bool HasTransitionTime => Month >= 1 && Month <= 12;

    /// <summary>
    ///     Gets or sets the delta.
    /// </summary>
    internal TimeSpan Delta
    {
        get => _delta;
        set => _delta = value;
    }

    /// <summary>
    ///     Gets or sets the time of day.
    /// </summary>
    internal TimeSpan TimeOfDay
    {
        get => _timeOfDay;
        set => _timeOfDay = value;
    }

    /// <summary>
    ///     Gets or sets a value that represents:
    ///     - The day of the month when Year is non zero,
    ///     - The index of the week in the month if Year is equal to zero.
    /// </summary>
    internal int DayOrder { get; set; }

    /// <summary>
    ///     Gets or sets the month.
    /// </summary>
    internal int Month { get; set; }

    /// <summary>
    ///     Gets or sets the day of the week.
    /// </summary>
    internal DayOfTheWeek DayOfTheWeek { get; set; }

    /// <summary>
    ///     Gets or sets the year. If Year is 0, the time change occurs every year according to a recurring pattern;
    ///     otherwise, the time change occurs at the date specified by Day, Month, Year.
    /// </summary>
    internal int Year { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="LegacyAvailabilityTimeZoneTime" /> class.
    /// </summary>
    internal LegacyAvailabilityTimeZoneTime()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="LegacyAvailabilityTimeZoneTime" /> class.
    /// </summary>
    /// <param name="transitionTime">The transition time used to initialize this instance.</param>
    /// <param name="delta">The offset used to initialize this instance.</param>
    internal LegacyAvailabilityTimeZoneTime(TransitionTime transitionTime, TimeSpan delta)
        : this()
    {
        _delta = delta;

        if (transitionTime.IsFixedDateRule)
        {
            // TimeZoneInfo doesn't support an actual year. Fixed date transitions occur at the same
            // date every year the adjustment rule the transition belongs to applies. The best thing
            // we can do here is use the current year.
            Year = DateTime.Today.Year;
            Month = transitionTime.Month;
            DayOrder = transitionTime.Day;
            _timeOfDay = transitionTime.TimeOfDay.TimeOfDay;
        }
        else
        {
            // For floating rules, the mapping is direct.
            Year = 0;
            Month = transitionTime.Month;
            DayOfTheWeek = EwsUtilities.SystemToEwsDayOfTheWeek(transitionTime.DayOfWeek);
            DayOrder = transitionTime.Week;
            _timeOfDay = transitionTime.TimeOfDay.TimeOfDay;
        }
    }

    /// <summary>
    ///     Converts this instance to TimeZoneInfo.TransitionTime.
    /// </summary>
    /// <returns>A TimeZoneInfo.TransitionTime</returns>
    internal TransitionTime ToTransitionTime()
    {
        if (Year == 0)
        {
            return TransitionTime.CreateFloatingDateRule(
                new DateTime(
                    DateTime.MinValue.Year,
                    DateTime.MinValue.Month,
                    DateTime.MinValue.Day,
                    _timeOfDay.Hours,
                    _timeOfDay.Minutes,
                    _timeOfDay.Seconds
                ),
                Month,
                DayOrder,
                EwsUtilities.EwsToSystemDayOfWeek(DayOfTheWeek)
            );
        }

        return TransitionTime.CreateFixedDateRule(new DateTime(_timeOfDay.Ticks), Month, DayOrder);
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
            case XmlElementNames.Bias:
            {
                _delta = TimeSpan.FromMinutes(reader.ReadElementValue<int>());
                return true;
            }
            case XmlElementNames.Time:
            {
                _timeOfDay = TimeSpan.Parse(reader.ReadElementValue());
                return true;
            }
            case XmlElementNames.DayOrder:
            {
                DayOrder = reader.ReadElementValue<int>();
                return true;
            }
            case XmlElementNames.DayOfWeek:
            {
                DayOfTheWeek = reader.ReadElementValue<DayOfTheWeek>();
                return true;
            }
            case XmlElementNames.Month:
            {
                Month = reader.ReadElementValue<int>();
                return true;
            }
            case XmlElementNames.Year:
            {
                Year = reader.ReadElementValue<int>();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes the elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Bias, (int)_delta.TotalMinutes);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Time, EwsUtilities.TimeSpanToXsTime(_timeOfDay));

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DayOrder, DayOrder);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Month, Month);

        // Only write DayOfWeek if this is a recurring time change
        if (Year == 0)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DayOfWeek, DayOfTheWeek);
        }

        // Only emit year if it's non zero, otherwise AS returns "Request is invalid"
        if (Year != 0)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Year, Year);
        }
    }
}
