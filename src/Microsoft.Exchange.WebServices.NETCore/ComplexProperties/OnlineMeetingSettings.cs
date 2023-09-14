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
///     Online Meeting Lobby Bypass options.
/// </summary>
[PublicAPI]
public enum LobbyBypass
{
    /// <summary>
    ///     Disabled.
    /// </summary>
    Disabled,

    /// <summary>
    ///     Enabled for gateway participants.
    /// </summary>
    EnabledForGatewayParticipants,
}

/// <summary>
///     Online Meeting Access Level options.
/// </summary>
[PublicAPI]
public enum OnlineMeetingAccessLevel
{
    /// <summary>
    ///     Locked.
    /// </summary>
    Locked,

    /// <summary>
    ///     Invited.
    /// </summary>
    Invited,

    /// <summary>
    ///     Internal.
    /// </summary>
    Internal,

    /// <summary>
    ///     Everyone.
    /// </summary>
    Everyone,
}

/// <summary>
///     Online Meeting Presenters options.
/// </summary>
[PublicAPI]
public enum Presenters
{
    /// <summary>
    ///     Disabled.
    /// </summary>
    Disabled,

    /// <summary>
    ///     Internal.
    /// </summary>
    Internal,

    /// <summary>
    ///     Everyone.
    /// </summary>
    Everyone,
}

/// <summary>
///     Represents Lync online meeting settings.
/// </summary>
[PublicAPI]
public class OnlineMeetingSettings : ComplexProperty
{
    /// <summary>
    ///     Email address.
    /// </summary>
    private LobbyBypass _lobbyBypass;

    /// <summary>
    ///     Routing type.
    /// </summary>
    private OnlineMeetingAccessLevel _accessLevel;

    /// <summary>
    ///     Routing type.
    /// </summary>
    private Presenters _presenters;

    /// <summary>
    ///     Initializes a new instance of the <see cref="OnlineMeetingSettings" /> class.
    /// </summary>
    public OnlineMeetingSettings()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="OnlineMeetingSettings" /> class.
    /// </summary>
    /// <param name="lobbyBypass">The address used to initialize the OnlineMeetingSettings.</param>
    /// <param name="accessLevel">The routing type used to initialize the OnlineMeetingSettings.</param>
    /// <param name="presenters">Mailbox type of the participant.</param>
    internal OnlineMeetingSettings(LobbyBypass lobbyBypass, OnlineMeetingAccessLevel accessLevel, Presenters presenters)
    {
        _lobbyBypass = lobbyBypass;
        _accessLevel = accessLevel;
        _presenters = presenters;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="OnlineMeetingSettings" /> class from another OnlineMeetingSettings
    ///     instance.
    /// </summary>
    /// <param name="onlineMeetingSettings">OnlineMeetingSettings instance to copy.</param>
    internal OnlineMeetingSettings(OnlineMeetingSettings onlineMeetingSettings)
        : this()
    {
        EwsUtilities.ValidateParam(onlineMeetingSettings, "OnlineMeetingSettings");

        LobbyBypass = onlineMeetingSettings.LobbyBypass;
        AccessLevel = onlineMeetingSettings.AccessLevel;
        Presenters = onlineMeetingSettings.Presenters;
    }

    /// <summary>
    ///     Gets or sets the online meeting setting that describes whether users dialing in by phone have to wait in the lobby.
    /// </summary>
    public LobbyBypass LobbyBypass
    {
        get => _lobbyBypass;
        set => SetFieldValue(ref _lobbyBypass, value);
    }

    /// <summary>
    ///     Gets or sets the online meeting setting that describes access permission to the meeting.
    /// </summary>
    public OnlineMeetingAccessLevel AccessLevel
    {
        get => _accessLevel;
        set => SetFieldValue(ref _accessLevel, value);
    }

    /// <summary>
    ///     Gets or sets the online meeting setting that defines the meeting leaders.
    /// </summary>
    public Presenters Presenters
    {
        get => _presenters;
        set => SetFieldValue(ref _presenters, value);
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
            case XmlElementNames.LobbyBypass:
            {
                _lobbyBypass = reader.ReadElementValue<LobbyBypass>();
                return true;
            }
            case XmlElementNames.AccessLevel:
            {
                _accessLevel = reader.ReadElementValue<OnlineMeetingAccessLevel>();
                return true;
            }
            case XmlElementNames.Presenters:
            {
                _presenters = reader.ReadElementValue<Presenters>();
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
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LobbyBypass, LobbyBypass);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AccessLevel, AccessLevel);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Presenters, Presenters);
    }
}
