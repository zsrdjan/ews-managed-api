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
using System.Xml;

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Represents an Outlook configuration settings account.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
internal sealed class OutlookAccount
{
    #region Private constants

    private const string Settings = "settings";
    private const string RedirectAddr = "redirectAddr";
    private const string RedirectUrl = "redirectUrl";

    #endregion


    #region Private fields

    private readonly Dictionary<OutlookProtocolType, OutlookProtocol> _protocols;
    private readonly AlternateMailboxCollection _alternateMailboxes;

    #endregion


    /// <summary>
    ///     Initializes a new instance of the <see cref="OutlookAccount" /> class.
    /// </summary>
    internal OutlookAccount()
    {
        _protocols = new Dictionary<OutlookProtocolType, OutlookProtocol>();
        _alternateMailboxes = new AlternateMailboxCollection();
    }

    /// <summary>
    ///     Load from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsXmlReader reader)
    {
        do
        {
            reader.Read();

            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.AccountType:
                    {
                        AccountType = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.Action:
                    {
                        var xmlResponseType = reader.ReadElementValue();

                        ResponseType = xmlResponseType switch
                        {
                            Settings => AutodiscoverResponseType.Success,
                            RedirectUrl => AutodiscoverResponseType.RedirectUrl,
                            RedirectAddr => AutodiscoverResponseType.RedirectAddress,
                            _ => AutodiscoverResponseType.Error,
                        };

                        break;
                    }
                    case XmlElementNames.Protocol:
                    {
                        var protocol = new OutlookProtocol();
                        protocol.LoadFromXml(reader);
                        if (_protocols.ContainsKey(protocol.ProtocolType))
                        {
                            // There should be strictly one node per protocol type in the autodiscover response.
                            throw new ServiceLocalException(Strings.InvalidAutodiscoverServiceResponse);
                        }

                        _protocols.Add(protocol.ProtocolType, protocol);
                        break;
                    }
                    case XmlElementNames.RedirectAddr:
                    case XmlElementNames.RedirectUrl:
                    {
                        RedirectTarget = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.AlternateMailboxes:
                    {
                        var alternateMailbox = AlternateMailbox.LoadFromXml(reader);
                        _alternateMailboxes.Entries.Add(alternateMailbox);
                        break;
                    }

                    default:
                    {
                        reader.SkipCurrentElement();
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.Account));
    }

    /// <summary>
    ///     Convert OutlookAccount to GetUserSettings response.
    /// </summary>
    /// <param name="requestedSettings">The requested settings.</param>
    /// <param name="response">GetUserSettings response.</param>
    internal void ConvertToUserSettings(List<UserSettingName> requestedSettings, GetUserSettingsResponse response)
    {
        foreach (var protocol in _protocols.Values)
        {
            protocol.ConvertToUserSettings(requestedSettings, response);
        }

        if (requestedSettings.Contains(UserSettingName.AlternateMailboxes))
        {
            response.Settings[UserSettingName.AlternateMailboxes] = _alternateMailboxes;
        }
    }

    /// <summary>
    ///     Gets or sets type of the account.
    /// </summary>
    internal string AccountType { get; set; }

    /// <summary>
    ///     Gets or sets the type of the response.
    /// </summary>
    internal AutodiscoverResponseType ResponseType { get; set; }

    /// <summary>
    ///     Gets or sets the redirect target.
    /// </summary>
    internal string RedirectTarget { get; set; }
}
