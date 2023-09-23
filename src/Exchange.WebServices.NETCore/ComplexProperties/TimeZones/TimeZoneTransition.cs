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
///     Represents the base class for all time zone transitions.
/// </summary>
internal class TimeZoneTransition : ComplexProperty
{
    private const string PeriodTarget = "Period";
    private const string GroupTarget = "Group";

    private readonly TimeZoneDefinition _timeZoneDefinition;
    private TimeZoneTransitionGroup _targetGroup;
    private TimeZonePeriod? _targetPeriod;

    /// <summary>
    ///     Gets the target period of the transition.
    /// </summary>
    internal TimeZonePeriod TargetPeriod => _targetPeriod;

    /// <summary>
    ///     Gets the target transition group of the transition.
    /// </summary>
    internal TimeZoneTransitionGroup TargetGroup => _targetGroup;

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZoneTransition" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition the transition will belong to.</param>
    internal TimeZoneTransition(TimeZoneDefinition timeZoneDefinition)
    {
        _timeZoneDefinition = timeZoneDefinition;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZoneTransition" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition the transition will belong to.</param>
    /// <param name="targetGroup">The transition group the transition will target.</param>
    internal TimeZoneTransition(TimeZoneDefinition timeZoneDefinition, TimeZoneTransitionGroup targetGroup)
        : this(timeZoneDefinition)
    {
        _targetGroup = targetGroup;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZoneTransition" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition the transition will belong to.</param>
    /// <param name="targetPeriod">The period the transition will target.</param>
    internal TimeZoneTransition(TimeZoneDefinition timeZoneDefinition, TimeZonePeriod targetPeriod)
        : this(timeZoneDefinition)
    {
        _targetPeriod = targetPeriod;
    }

    /// <summary>
    ///     Creates a time zone period transition of the appropriate type given an XML element name.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition to which the transition will belong.</param>
    /// <param name="xmlElementName">The XML element name.</param>
    /// <returns>A TimeZonePeriodTransition instance.</returns>
    internal static TimeZoneTransition Create(TimeZoneDefinition timeZoneDefinition, string xmlElementName)
    {
        return xmlElementName switch
        {
            XmlElementNames.AbsoluteDateTransition => new AbsoluteDateTransition(timeZoneDefinition),
            XmlElementNames.RecurringDayTransition => new RelativeDayOfMonthTransition(timeZoneDefinition),
            XmlElementNames.RecurringDateTransition => new AbsoluteDayOfMonthTransition(timeZoneDefinition),
            XmlElementNames.Transition => new TimeZoneTransition(timeZoneDefinition),
            _ => throw new ServiceLocalException(
                string.Format(Strings.UnknownTimeZonePeriodTransitionType, xmlElementName)
            ),
        };
    }

    /// <summary>
    ///     Creates a time zone transition based on the specified transition time.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition that will own the transition.</param>
    /// <param name="targetPeriod">The period the transition will target.</param>
    /// <param name="transitionTime">The transition time to initialize from.</param>
    /// <returns>A TimeZoneTransition.</returns>
    internal static TimeZoneTransition CreateTimeZoneTransition(
        TimeZoneDefinition timeZoneDefinition,
        TimeZonePeriod targetPeriod,
        TransitionTime transitionTime
    )
    {
        TimeZoneTransition transition;

        if (transitionTime.IsFixedDateRule)
        {
            transition = new AbsoluteDayOfMonthTransition(timeZoneDefinition, targetPeriod);
        }
        else
        {
            transition = new RelativeDayOfMonthTransition(timeZoneDefinition, targetPeriod);
        }

        transition.InitializeFromTransitionTime(transitionTime);

        return transition;
    }

    /// <summary>
    ///     Gets the XML element name associated with the transition.
    /// </summary>
    /// <returns>The XML element name associated with the transition.</returns>
    internal virtual string GetXmlElementName()
    {
        return XmlElementNames.Transition;
    }

    /// <summary>
    ///     Creates a time zone transition time.
    /// </summary>
    /// <returns>A TimeZoneInfo.TransitionTime.</returns>
    internal virtual TransitionTime CreateTransitionTime()
    {
        throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
    }

    /// <summary>
    ///     Initializes this transition based on the specified transition time.
    /// </summary>
    /// <param name="transitionTime">The transition time to initialize from.</param>
    internal virtual void InitializeFromTransitionTime(TransitionTime transitionTime)
    {
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
            case XmlElementNames.To:
            {
                var targetKind = reader.ReadAttributeValue(XmlAttributeNames.Kind);
                var targetId = reader.ReadElementValue();

                switch (targetKind)
                {
                    case PeriodTarget:
                    {
                        if (!_timeZoneDefinition.Periods.TryGetValue(targetId, out _targetPeriod))
                        {
                            throw new ServiceLocalException(string.Format(Strings.PeriodNotFound, targetId));
                        }

                        break;
                    }
                    case GroupTarget:
                    {
                        if (!_timeZoneDefinition.TransitionGroups.TryGetValue(targetId, out _targetGroup))
                        {
                            throw new ServiceLocalException(string.Format(Strings.TransitionGroupNotFound, targetId));
                        }

                        break;
                    }
                    default:
                    {
                        throw new ServiceLocalException(Strings.UnsupportedTimeZonePeriodTransitionTarget);
                    }
                }

                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.To);

        if (_targetPeriod != null)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Kind, PeriodTarget);
            writer.WriteValue(_targetPeriod.Id, XmlElementNames.To);
        }
        else
        {
            writer.WriteAttributeValue(XmlAttributeNames.Kind, GroupTarget);
            writer.WriteValue(_targetGroup.Id, XmlElementNames.To);
        }

        writer.WriteEndElement(); // To
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        LoadFromXml(reader, GetXmlElementName());
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        WriteToXml(writer, GetXmlElementName());
    }
}
