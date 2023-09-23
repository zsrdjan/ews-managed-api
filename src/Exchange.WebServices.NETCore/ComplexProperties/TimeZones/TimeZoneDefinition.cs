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

using Microsoft.Exchange.WebServices.Data.Misc;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a time zone as defined by the EWS schema.
/// </summary>
[PublicAPI]
public class TimeZoneDefinition : ComplexProperty
{
    /// <summary>
    ///     Prefix for generated ids.
    /// </summary>
    private const string NoIdPrefix = "NoId_";

    private readonly List<TimeZoneTransition> _transitions = new();

    /// <summary>
    ///     Gets or sets the name of this time zone definition.
    /// </summary>
    internal string Name { get; set; }

    /// <summary>
    ///     Gets or sets the Id of this time zone definition.
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    ///     Gets the periods associated with this time zone definition, indexed by Id.
    /// </summary>
    internal Dictionary<string, TimeZonePeriod> Periods { get; } = new();

    /// <summary>
    ///     Gets the transition groups associated with this time zone definition, indexed by Id.
    /// </summary>
    internal Dictionary<string, TimeZoneTransitionGroup> TransitionGroups { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZoneDefinition" /> class.
    /// </summary>
    internal TimeZoneDefinition()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZoneDefinition" /> class.
    /// </summary>
    /// <param name="timeZoneInfo">The time zone info used to initialize this definition.</param>
    internal TimeZoneDefinition(TimeZoneInfo timeZoneInfo)
        : this()
    {
        Id = timeZoneInfo.Id;
        Name = timeZoneInfo.DisplayName;

        // TimeZoneInfo only supports one standard period, which bias is the time zone's base
        // offset to UTC.
        var standardPeriod = new TimeZonePeriod
        {
            Id = TimeZonePeriod.StandardPeriodId,
            Name = TimeZonePeriod.StandardPeriodName,
            Bias = -timeZoneInfo.BaseUtcOffset,
        };

        var adjustmentRules = timeZoneInfo.GetAdjustmentRulesEx();

        var transitionToStandardPeriod = new TimeZoneTransition(this, standardPeriod);

        if (adjustmentRules.Length == 0)
        {
            Periods.Add(standardPeriod.Id, standardPeriod);

            // If the time zone info doesn't support Daylight Saving Time, we just need to
            // create one transition to one group with one transition to the standard period.
            var transitionGroup = new TimeZoneTransitionGroup(this, "0");
            transitionGroup.Transitions.Add(transitionToStandardPeriod);

            TransitionGroups.Add(transitionGroup.Id, transitionGroup);

            var initialTransition = new TimeZoneTransition(this, transitionGroup);

            _transitions.Add(initialTransition);
        }
        else
        {
            for (var i = 0; i < adjustmentRules.Length; i++)
            {
                var transitionGroup = new TimeZoneTransitionGroup(this, TransitionGroups.Count.ToString());
                transitionGroup.InitializeFromAdjustmentRule(adjustmentRules[i], standardPeriod);

                TransitionGroups.Add(transitionGroup.Id, transitionGroup);

                TimeZoneTransition transition;

                if (i == 0)
                {
                    // If the first adjustment rule's start date in not undefined (DateTime.MinValue)
                    // we need to add a dummy group with a single, simple transition to the Standard
                    // period and a group containing the transitions mapping to the adjustment rule.
                    if (adjustmentRules[i].DateStart > DateTime.MinValue.Date)
                    {
                        var transitionToDummyGroup = new TimeZoneTransition(
                            this,
                            CreateTransitionGroupToPeriod(standardPeriod)
                        );

                        _transitions.Add(transitionToDummyGroup);

                        var absoluteDateTransition = new AbsoluteDateTransition(this, transitionGroup)
                        {
                            DateTime = adjustmentRules[i].DateStart,
                        };

                        transition = absoluteDateTransition;
                        Periods.Add(standardPeriod.Id, standardPeriod);
                    }
                    else
                    {
                        transition = new TimeZoneTransition(this, transitionGroup);
                    }
                }
                else
                {
                    var absoluteDateTransition = new AbsoluteDateTransition(this, transitionGroup)
                    {
                        DateTime = adjustmentRules[i].DateStart,
                    };

                    transition = absoluteDateTransition;
                }

                _transitions.Add(transition);
            }

            // If the last adjustment rule's end date is not undefined (DateTime.MaxValue),
            // we need to create another absolute date transition that occurs the date after
            // the last rule's end date. We target this additional transition to a group that
            // contains a single simple transition to the Standard period.
            var lastAdjustmentRuleEndDate = adjustmentRules[adjustmentRules.Length - 1].DateEnd;

            if (lastAdjustmentRuleEndDate < DateTime.MaxValue.Date)
            {
                var transitionToDummyGroup = new AbsoluteDateTransition(
                    this,
                    CreateTransitionGroupToPeriod(standardPeriod)
                )
                {
                    DateTime = lastAdjustmentRuleEndDate.AddDays(1),
                };

                _transitions.Add(transitionToDummyGroup);
            }
        }
    }

    /// <summary>
    ///     Compares the transitions.
    /// </summary>
    /// <param name="x">The first transition.</param>
    /// <param name="y">The second transition.</param>
    /// <returns>A negative number if x is less than y, 0 if x and y are equal, a positive number if x is greater than y.</returns>
    private int CompareTransitions(TimeZoneTransition x, TimeZoneTransition y)
    {
        if (x == y)
        {
            return 0;
        }

        if (x.GetType() == typeof(TimeZoneTransition))
        {
            return -1;
        }

        if (y.GetType() == typeof(TimeZoneTransition))
        {
            return 1;
        }

        var firstTransition = (AbsoluteDateTransition)x;
        var secondTransition = (AbsoluteDateTransition)y;

        return DateTime.Compare(firstTransition.DateTime, secondTransition.DateTime);
    }

    /// <summary>
    ///     Adds a transition group with a single transition to the specified period.
    /// </summary>
    /// <param name="timeZonePeriod">The time zone period.</param>
    /// <returns>A TimeZoneTransitionGroup.</returns>
    private TimeZoneTransitionGroup CreateTransitionGroupToPeriod(TimeZonePeriod timeZonePeriod)
    {
        var transitionToPeriod = new TimeZoneTransition(this, timeZonePeriod);

        var transitionGroup = new TimeZoneTransitionGroup(this, TransitionGroups.Count.ToString());
        transitionGroup.Transitions.Add(transitionToPeriod);

        TransitionGroups.Add(transitionGroup.Id, transitionGroup);

        return transitionGroup;
    }

    /// <summary>
    ///     Reads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        Name = reader.ReadAttributeValue(XmlAttributeNames.Name);
        Id = reader.ReadAttributeValue(XmlAttributeNames.Id);

        // EWS can return a TimeZone definition with no Id. Generate a new Id in this case.
        if (string.IsNullOrEmpty(Id))
        {
            var nameValue = string.IsNullOrEmpty(Name) ? string.Empty : Name;
            Id = NoIdPrefix + Math.Abs(nameValue.GetHashCode());
        }
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        // The Name attribute is only supported in Exchange 2010 and above.
        if (writer.Service.RequestedServerVersion != ExchangeVersion.Exchange2007_SP1)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Name, Name);
        }

        writer.WriteAttributeValue(XmlAttributeNames.Id, Id);
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
            case XmlElementNames.Periods:
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Period))
                    {
                        var period = new TimeZonePeriod();
                        period.LoadFromXml(reader);

                        // OM:1648848 Bad timezone data from clients can include duplicate rules
                        // for one year, with duplicate ID. In that case, let the first one win.
                        if (!Periods.ContainsKey(period.Id))
                        {
                            Periods.Add(period.Id, period);
                        }
                        else
                        {
                            reader.Service.TraceMessage(
                                TraceFlags.EwsTimeZones,
                                string.Format(
                                    "An entry with the same key (Id) '{0}' already exists in Periods. Cannot add another one. Existing entry: [Name='{1}', Bias='{2}']. Entry to skip: [Name='{3}', Bias='{4}'].",
                                    period.Id,
                                    Periods[period.Id].Name,
                                    Periods[period.Id].Bias,
                                    period.Name,
                                    period.Bias
                                )
                            );
                        }
                    }
                } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Periods));

                return true;
            }
            case XmlElementNames.TransitionsGroups:
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.TransitionsGroup))
                    {
                        var transitionGroup = new TimeZoneTransitionGroup(this);

                        transitionGroup.LoadFromXml(reader);

                        TransitionGroups.Add(transitionGroup.Id, transitionGroup);
                    }
                } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.TransitionsGroups));

                return true;
            }
            case XmlElementNames.Transitions:
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement())
                    {
                        var transition = TimeZoneTransition.Create(this, reader.LocalName);

                        transition.LoadFromXml(reader);

                        _transitions.Add(transition);
                    }
                } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Transitions));

                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        LoadFromXml(reader, XmlElementNames.TimeZoneDefinition);

        _transitions.Sort(CompareTransitions);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        // We only emit the full time zone definition against Exchange 2010 servers and above.
        if (writer.Service.RequestedServerVersion != ExchangeVersion.Exchange2007_SP1)
        {
            if (Periods.Count > 0)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Periods);

                foreach (var keyValuePair in Periods)
                {
                    keyValuePair.Value.WriteToXml(writer);
                }

                writer.WriteEndElement(); // Periods
            }

            if (TransitionGroups.Count > 0)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.TransitionsGroups);

                foreach (var keyValuePair in TransitionGroups)
                {
                    keyValuePair.Value.WriteToXml(writer);
                }

                writer.WriteEndElement(); // TransitionGroups
            }

            if (_transitions.Count > 0)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Transitions);

                foreach (var transition in _transitions)
                {
                    transition.WriteToXml(writer);
                }

                writer.WriteEndElement(); // Transitions
            }
        }
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        WriteToXml(writer, XmlElementNames.TimeZoneDefinition);
    }

    /// <summary>
    ///     Validates this time zone definition.
    /// </summary>
    internal void Validate()
    {
        // The definition must have at least one period, one transition group and one transition,
        // and there must be as many transitions as there are transition groups.
        if (Periods.Count < 1 ||
            _transitions.Count < 1 ||
            TransitionGroups.Count < 1 ||
            TransitionGroups.Count != _transitions.Count)
        {
            throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
        }

        // The first transition must be of type TimeZoneTransition.
        if (_transitions[0].GetType() != typeof(TimeZoneTransition))
        {
            throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
        }

        // All transitions must be to transition groups and be either TimeZoneTransition or
        // AbsoluteDateTransition instances.
        foreach (var transition in _transitions)
        {
            var transitionType = transition.GetType();

            if (transitionType != typeof(TimeZoneTransition) && transitionType != typeof(AbsoluteDateTransition))
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }

            if (transition.TargetGroup == null)
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }
        }

        // All transition groups must be valid.
        foreach (var transitionGroup in TransitionGroups.Values)
        {
            transitionGroup.Validate();
        }
    }

    /// <summary>
    ///     Converts this time zone definition into a TimeZoneInfo structure.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <returns>A TimeZoneInfo representing the same time zone as this definition.</returns>
    internal TimeZoneInfo ToTimeZoneInfo(ExchangeService service)
    {
        Validate();

        TimeZoneInfo result;

        // Retrieve the base offset to UTC, standard and daylight display names from
        // the last transition group, which is the one that currently applies given that
        // transitions are ordered chronologically.
        var creationParams = _transitions[_transitions.Count - 1].TargetGroup.GetCustomTimeZoneCreationParams();

        var adjustmentRules = new List<AdjustmentRule>();

        var startDate = DateTime.MinValue;

        for (var i = 0; i < _transitions.Count; i++)
        {
            DateTime endDate;
            DateTime effectiveEndDate;
            if (i < _transitions.Count - 1)
            {
                endDate = (_transitions[i + 1] as AbsoluteDateTransition).DateTime;
                effectiveEndDate = endDate.AddDays(-1);
            }
            else
            {
                endDate = DateTime.MaxValue;
                effectiveEndDate = endDate;
            }

            // OM:1648848 Due to bad timezone data from clients the 
            // startDate may not always come before the effectiveEndDate
            if (startDate < effectiveEndDate)
            {
                var adjustmentRule = _transitions[i].TargetGroup.CreateAdjustmentRule(startDate, effectiveEndDate);

                if (adjustmentRule != null)
                {
                    adjustmentRules.Add(adjustmentRule);
                }

                startDate = endDate;
            }
            else
            {
                service.TraceMessage(
                    TraceFlags.EwsTimeZones,
                    string.Format(
                        "The startDate '{0}' is not before the effectiveEndDate '{1}'. Will skip creating adjustment rule.",
                        startDate,
                        effectiveEndDate
                    )
                );
            }
        }

        if (adjustmentRules.Count == 0)
        {
            // If there are no adjustment rule, the time zone does not support Daylight
            // saving time.
            result = TimeZoneExtensions.CreateCustomTimeZone(
                Id,
                creationParams.BaseOffsetToUtc,
                Name,
                creationParams.StandardDisplayName
            );
        }
        else
        {
            result = TimeZoneExtensions.CreateCustomTimeZone(
                Id,
                creationParams.BaseOffsetToUtc,
                Name,
                creationParams.StandardDisplayName,
                creationParams.DaylightDisplayName,
                adjustmentRules.ToArray()
            );
        }

        return result;
    }
}
