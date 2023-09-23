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

using System.Collections.ObjectModel;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the availability of an individual attendee.
/// </summary>
[PublicAPI]
public sealed class AttendeeAvailability : ServiceResponse
{
    /// <summary>
    ///     Gets a collection of calendar events for the attendee.
    /// </summary>
    public Collection<CalendarEvent> CalendarEvents { get; } = new();

    /// <summary>
    ///     Gets the free/busy view type that wes retrieved for the attendee.
    /// </summary>
    public FreeBusyViewType ViewType { get; private set; }

    /// <summary>
    ///     Gets a collection of merged free/busy status for the attendee.
    /// </summary>
    public Collection<LegacyFreeBusyStatus> MergedFreeBusyStatus { get; } = new();

    /// <summary>
    ///     Gets the working hours of the attendee.
    /// </summary>
    public WorkingHours WorkingHours { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AttendeeAvailability" /> class.
    /// </summary>
    internal AttendeeAvailability()
    {
    }

    /// <summary>
    ///     Loads the free busy view from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="viewType">Type of free/busy view.</param>
    internal void LoadFreeBusyViewFromXml(EwsServiceXmlReader reader, FreeBusyViewType viewType)
    {
        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.FreeBusyView);

        var viewTypeString = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.FreeBusyViewType);

        ViewType = Enum.Parse<FreeBusyViewType>(viewTypeString, false);

        do
        {
            reader.Read();

            if (reader.IsStartElement())
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.MergedFreeBusy:
                    {
                        var mergedFreeBusy = reader.ReadElementValue();

                        for (var i = 0; i < mergedFreeBusy.Length; i++)
                        {
                            MergedFreeBusyStatus.Add((LegacyFreeBusyStatus)byte.Parse(mergedFreeBusy[i].ToString()));
                        }

                        break;
                    }
                    case XmlElementNames.CalendarEventArray:
                    {
                        do
                        {
                            reader.Read();

                            // Sometimes Exchange Online returns blank CalendarEventArray tag like bellow.
                            // <CalendarEventArray xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
                            // So we have to check the end of CalendarEventArray tag.
                            if (reader.LocalName == XmlElementNames.FreeBusyView)
                            {
                                // There is no the end tag of CalendarEventArray, but the reader is reading the end tag of FreeBusyView.
                                break;
                            }

                            if (reader.LocalName == XmlElementNames.WorkingHours)
                            {
                                // There is no the end tag of CalendarEventArray, but the reader is reading the start tag of WorkingHours.
                                goto case XmlElementNames.WorkingHours;
                            }

                            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.CalendarEvent))
                            {
                                var calendarEvent = new CalendarEvent();

                                calendarEvent.LoadFromXml(reader, XmlElementNames.CalendarEvent);

                                CalendarEvents.Add(calendarEvent);
                            }
                        } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.CalendarEventArray));

                        break;
                    }
                    case XmlElementNames.WorkingHours:
                    {
                        WorkingHours = new WorkingHours();
                        WorkingHours.LoadFromXml(reader, reader.LocalName);

                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.FreeBusyView));
    }
}
