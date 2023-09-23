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
///     Represents a group of time zone period transitions.
/// </summary>
internal sealed class TimeZoneTransitionGroup : ComplexProperty
{
    private readonly TimeZoneDefinition _timeZoneDefinition;
    private TimeZoneTransition _transitionToDaylight;
    private TimeZoneTransition _transitionToStandard;

    /// <summary>
    ///     Gets a value indicating whether this group contains a transition to the Daylight period.
    /// </summary>
    /// <value><c>true</c> if this group contains a transition to daylight; otherwise, <c>false</c>.</value>
    internal bool SupportsDaylight => Transitions.Count == 2;

    /// <summary>
    ///     Gets the transition to the Daylight period.
    /// </summary>
    private TimeZoneTransition TransitionToDaylight
    {
        get
        {
            InitializeTransitions();

            return _transitionToDaylight;
        }
    }

    /// <summary>
    ///     Gets the transition to the Standard period.
    /// </summary>
    private TimeZoneTransition TransitionToStandard
    {
        get
        {
            InitializeTransitions();

            return _transitionToStandard;
        }
    }

    /// <summary>
    ///     Gets or sets the id of this group.
    /// </summary>
    internal string Id { get; set; }

    /// <summary>
    ///     Gets the transitions in this group.
    /// </summary>
    internal List<TimeZoneTransition> Transitions { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZoneTransitionGroup" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition.</param>
    internal TimeZoneTransitionGroup(TimeZoneDefinition timeZoneDefinition)
    {
        _timeZoneDefinition = timeZoneDefinition;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZoneTransitionGroup" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition.</param>
    /// <param name="id">The Id of the new transition group.</param>
    internal TimeZoneTransitionGroup(TimeZoneDefinition timeZoneDefinition, string id)
        : this(timeZoneDefinition)
    {
        Id = id;
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        LoadFromXml(reader, XmlElementNames.TransitionsGroup);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        WriteToXml(writer, XmlElementNames.TransitionsGroup);
    }

    /// <summary>
    ///     Reads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        Id = reader.ReadAttributeValue(XmlAttributeNames.Id);
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.Id, Id);
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        reader.EnsureCurrentNodeIsStartElement();

        var transition = TimeZoneTransition.Create(_timeZoneDefinition, reader.LocalName);

        transition.LoadFromXml(reader);

        EwsUtilities.Assert(
            transition.TargetPeriod != null,
            "TimeZoneTransitionGroup.TryReadElementFromXml",
            "The transition's target period is null."
        );

        Transitions.Add(transition);

        return true;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        foreach (var transition in Transitions)
        {
            transition.WriteToXml(writer);
        }
    }

    /// <summary>
    ///     Initializes this transition group based on the specified adjustment rule.
    /// </summary>
    /// <param name="adjustmentRule">The adjustment rule to initialize from.</param>
    /// <param name="standardPeriod">A reference to the pre-created standard period.</param>
    internal void InitializeFromAdjustmentRule(AdjustmentRule adjustmentRule, TimeZonePeriod standardPeriod)
    {
        if (adjustmentRule.DaylightDelta.TotalSeconds == 0)
        {
            // If the time zone info doesn't support Daylight Saving Time, we just need to
            // create one transition to one group with one transition to the standard period.
            var standardPeriodToSet = new TimeZonePeriod
            {
                Id = $"{standardPeriod.Id}/{adjustmentRule.DateStart.Year}",
                Name = standardPeriod.Name,
                Bias = standardPeriod.Bias,
            };
            _timeZoneDefinition.Periods.AddOrUpdate(standardPeriodToSet.Id, standardPeriodToSet);

            _transitionToStandard = new TimeZoneTransition(_timeZoneDefinition, standardPeriodToSet);
            Transitions.Add(_transitionToStandard);
        }
        else
        {
            var daylightPeriod = new TimeZonePeriod
            {
                // Generate an Id of the form "Daylight/2008"
                Id = $"{TimeZonePeriod.DaylightPeriodId}/{adjustmentRule.DateStart.Year}",
                Name = TimeZonePeriod.DaylightPeriodName,
                Bias = standardPeriod.Bias - adjustmentRule.DaylightDelta,
            };

            _timeZoneDefinition.Periods.AddOrUpdate(daylightPeriod.Id, daylightPeriod);

            _transitionToDaylight = TimeZoneTransition.CreateTimeZoneTransition(
                _timeZoneDefinition,
                daylightPeriod,
                adjustmentRule.DaylightTransitionStart
            );

            var standardPeriodToSet = new TimeZonePeriod
            {
                Id = $"{standardPeriod.Id}/{adjustmentRule.DateStart.Year}",
                Name = standardPeriod.Name,
                Bias = standardPeriod.Bias,
            };
            _timeZoneDefinition.Periods.AddOrUpdate(standardPeriodToSet.Id, standardPeriodToSet);

            _transitionToStandard = TimeZoneTransition.CreateTimeZoneTransition(
                _timeZoneDefinition,
                standardPeriodToSet,
                adjustmentRule.DaylightTransitionEnd
            );

            Transitions.Add(_transitionToDaylight);
            Transitions.Add(_transitionToStandard);
        }
    }

    /// <summary>
    ///     Validates this transition group.
    /// </summary>
    internal void Validate()
    {
        // There must be exactly one or two transitions in the group.
        if (Transitions.Count < 1 || Transitions.Count > 2)
        {
            throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
        }

        // If there is only one transition, it must be of type TimeZoneTransition
        if (Transitions.Count == 1 && !(Transitions[0].GetType() == typeof(TimeZoneTransition)))
        {
            throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
        }

        // If there are two transitions, none of them should be of type TimeZoneTransition
        if (Transitions.Count == 2)
        {
            foreach (var transition in Transitions)
            {
                if (transition.GetType() == typeof(TimeZoneTransition))
                {
                    throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
                }
            }
        }

        // All the transitions in the group must be to a period.
        foreach (var transition in Transitions)
        {
            if (transition.TargetPeriod == null)
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }
        }
    }

    /// <summary>
    ///     Initializes the private members holding references to the transitions to the Daylight
    ///     and Standard periods.
    /// </summary>
    private void InitializeTransitions()
    {
        if (_transitionToStandard == null)
        {
            foreach (var transition in Transitions)
            {
                if (transition.TargetPeriod.IsStandardPeriod || Transitions.Count == 1)
                {
                    _transitionToStandard = transition;
                }
                else
                {
                    _transitionToDaylight = transition;
                }
            }
        }

        // If we didn't find a Standard period, this is an invalid time zone group.
        if (_transitionToStandard == null)
        {
            throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
        }
    }

    /// <summary>
    ///     Gets the offset to UTC based on this group's transitions.
    /// </summary>
    internal CustomTimeZoneCreateParams GetCustomTimeZoneCreationParams()
    {
        var result = new CustomTimeZoneCreateParams();

        if (TransitionToDaylight != null)
        {
            result.DaylightDisplayName = TransitionToDaylight.TargetPeriod.Name;
        }

        result.StandardDisplayName = TransitionToStandard.TargetPeriod.Name;

        // Assume that the standard period's offset is the base offset to UTC.
        // EWS returns a positive offset for time zones that are behind UTC, and
        // a negative one for time zones ahead of UTC. TimeZoneInfo does it the other
        // way around.
        result.BaseOffsetToUtc = -TransitionToStandard.TargetPeriod.Bias;

        return result;
    }

    /// <summary>
    ///     Gets the delta offset for the daylight.
    /// </summary>
    /// <returns></returns>
    internal TimeSpan GetDaylightDelta()
    {
        if (SupportsDaylight)
        {
            // EWS returns a positive offset for time zones that are behind UTC, and
            // a negative one for time zones ahead of UTC. TimeZoneInfo does it the other
            // way around.
            return TransitionToStandard.TargetPeriod.Bias - TransitionToDaylight.TargetPeriod.Bias;
        }

        return TimeSpan.Zero;
    }

    /// <summary>
    ///     Creates a time zone adjustment rule.
    /// </summary>
    /// <param name="startDate">The start date of the adjustment rule.</param>
    /// <param name="endDate">The end date of the adjustment rule.</param>
    /// <returns>An TimeZoneInfo.AdjustmentRule.</returns>
    internal AdjustmentRule? CreateAdjustmentRule(DateTime startDate, DateTime endDate)
    {
        // If there is only one transition, we can't create an adjustment rule. We have to assume
        // that the base offset to UTC is unchanged.
        if (Transitions.Count == 1)
        {
            return null;
        }

        return AdjustmentRule.CreateAdjustmentRule(
            startDate.Date,
            endDate.Date,
            GetDaylightDelta(),
            TransitionToDaylight.CreateTransitionTime(),
            TransitionToStandard.CreateTransitionTime()
        );
    }

    /// <summary>
    ///     Represents custom time zone creation parameters.
    /// </summary>
    internal class CustomTimeZoneCreateParams
    {
        /// <summary>
        ///     Gets or sets the base offset to UTC.
        /// </summary>
        internal TimeSpan BaseOffsetToUtc { get; set; }

        /// <summary>
        ///     Gets or sets the display name of the standard period.
        /// </summary>
        internal string StandardDisplayName { get; set; }

        /// <summary>
        ///     Gets or sets the display name of the daylight period.
        /// </summary>
        internal string DaylightDisplayName { get; set; }

        /// <summary>
        ///     Gets a value indicating whether the custom time zone should have a daylight period.
        /// </summary>
        /// <value>
        ///     <c>true</c> if the custom time zone should have a daylight period; otherwise, <c>false</c>.
        /// </value>
        internal bool HasDaylightPeriod => !string.IsNullOrEmpty(DaylightDisplayName);

        /// <summary>
        ///     Initializes a new instance of the <see cref="CustomTimeZoneCreateParams" /> class.
        /// </summary>
        internal CustomTimeZoneCreateParams()
        {
        }
    }
}
