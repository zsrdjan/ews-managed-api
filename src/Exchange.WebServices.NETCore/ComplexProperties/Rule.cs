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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a rule that automatically handles incoming messages.
///     A rule consists of a set of conditions and exceptions that determine whether or
///     not a set of actions should be executed on incoming messages.
/// </summary>
[PublicAPI]
public sealed class Rule : ComplexProperty
{
    /// <summary>
    ///     The rule display name.
    /// </summary>
    private string _displayName;

    /// <summary>
    ///     The rule status of enabled or not.
    /// </summary>
    private bool _isEnabled;

    /// <summary>
    ///     The rule status of in error or not.
    /// </summary>
    private bool _isInError;

    /// <summary>
    ///     The rule priority.
    /// </summary>
    private int _priority;

    /// <summary>
    ///     The rule ID.
    /// </summary>
    private string _ruleId;

    /// <summary>
    ///     Gets or sets the Id of this rule.
    /// </summary>
    public string Id
    {
        get => _ruleId;
        set => SetFieldValue(ref _ruleId, value);
    }

    /// <summary>
    ///     Gets or sets the name of this rule as it should be displayed to the user.
    /// </summary>
    public string DisplayName
    {
        get => _displayName;
        set => SetFieldValue(ref _displayName, value);
    }

    /// <summary>
    ///     Gets or sets the priority of this rule, which determines its execution order.
    /// </summary>
    public int Priority
    {
        get => _priority;
        set => SetFieldValue(ref _priority, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating whether this rule is enabled.
    /// </summary>
    public bool IsEnabled
    {
        get => _isEnabled;
        set => SetFieldValue(ref _isEnabled, value);
    }

    /// <summary>
    ///     Gets a value indicating whether this rule can be modified via EWS.
    ///     If IsNotSupported is true, the rule cannot be modified via EWS.
    /// </summary>
    public bool IsNotSupported { get; private set; }

    /// <summary>
    ///     Gets or sets a value indicating whether this rule has errors. A rule that is in error
    ///     cannot be processed unless it is updated and the error is corrected.
    /// </summary>
    public bool IsInError
    {
        get => _isInError;
        set => SetFieldValue(ref _isInError, value);
    }

    /// <summary>
    ///     Gets the conditions that determine whether or not this rule should be
    ///     executed against incoming messages.
    /// </summary>
    public RulePredicates Conditions { get; }

    /// <summary>
    ///     Gets the actions that should be executed against incoming messages if the
    ///     conditions evaluate as true.
    /// </summary>
    public RuleActions Actions { get; }

    /// <summary>
    ///     Gets the exceptions that determine if this rule should be skipped even if
    ///     its conditions evaluate to true.
    /// </summary>
    public RulePredicates Exceptions { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Rule" /> class.
    /// </summary>
    public Rule()
    {
        //// New rule has priority as 0 by default
        _priority = 1;
        //// New rule is enabled by default
        _isEnabled = true;
        Conditions = new RulePredicates();
        Actions = new RuleActions();
        Exceptions = new RulePredicates();
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
            case XmlElementNames.DisplayName:
            {
                _displayName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.RuleId:
            {
                _ruleId = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.Priority:
            {
                _priority = reader.ReadElementValue<int>();
                return true;
            }
            case XmlElementNames.IsEnabled:
            {
                _isEnabled = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsNotSupported:
            {
                IsNotSupported = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.IsInError:
            {
                _isInError = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.Conditions:
            {
                Conditions.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.Actions:
            {
                Actions.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.Exceptions:
            {
                Exceptions.LoadFromXml(reader, reader.LocalName);
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
        if (!string.IsNullOrEmpty(Id))
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RuleId, Id);
        }

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DisplayName, DisplayName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Priority, Priority);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsEnabled, IsEnabled);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsInError, IsInError);
        Conditions.WriteToXml(writer, XmlElementNames.Conditions);
        Exceptions.WriteToXml(writer, XmlElementNames.Exceptions);
        Actions.WriteToXml(writer, XmlElementNames.Actions);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void InternalValidate()
    {
        base.InternalValidate();
        EwsUtilities.ValidateParam(_displayName, "DisplayName");
        EwsUtilities.ValidateParam(Conditions);
        EwsUtilities.ValidateParam(Exceptions);
        EwsUtilities.ValidateParam(Actions);
    }
}
