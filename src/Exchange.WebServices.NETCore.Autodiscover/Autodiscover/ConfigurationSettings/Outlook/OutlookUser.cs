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
///     Represents the user Outlook configuration settings apply to.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
internal sealed class OutlookUser
{
    /// <summary>
    ///     Converters to translate Outlook user settings.
    ///     Each entry maps to a lambda expression used to get the matching property from the OutlookUser instance.
    /// </summary>
    private static readonly IReadOnlyDictionary<UserSettingName, Func<OutlookUser, string>> ConverterDictionary =
        new Dictionary<UserSettingName, Func<OutlookUser, string>>
        {
            // @formatter:off
            { UserSettingName.UserDisplayName, u => u._displayName },
            { UserSettingName.UserDN, u => u._legacyDn },
            { UserSettingName.UserDeploymentId, u => u._deploymentId },
            { UserSettingName.AutoDiscoverSMTPAddress, u => u._autodiscoverAmtpAddress },
            // @formatter:on
        };

    private string _displayName;
    private string _legacyDn;
    private string _deploymentId;
    private string _autodiscoverAmtpAddress;


    /// <summary>
    ///     Initializes a new instance of the <see cref="OutlookUser" /> class.
    /// </summary>
    internal OutlookUser()
    {
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
                    case XmlElementNames.DisplayName:
                    {
                        _displayName = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.LegacyDN:
                    {
                        _legacyDn = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.DeploymentId:
                    {
                        _deploymentId = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.AutoDiscoverSMTPAddress:
                    {
                        _autodiscoverAmtpAddress = reader.ReadElementValue();
                        break;
                    }
                    default:
                    {
                        reader.SkipCurrentElement();
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.User));
    }

    /// <summary>
    ///     Convert OutlookUser to GetUserSettings response.
    /// </summary>
    /// <param name="requestedSettings">The requested settings.</param>
    /// <param name="response">The response.</param>
    internal void ConvertToUserSettings(List<UserSettingName> requestedSettings, GetUserSettingsResponse response)
    {
        // In English: collect converters that are contained in the requested settings.
        var converterQuery = from converter in ConverterDictionary
            where requestedSettings.Contains(converter.Key)
            select converter;

        foreach (var kv in converterQuery)
        {
            var value = kv.Value(this);
            if (!string.IsNullOrEmpty(value))
            {
                response.Settings[kv.Key] = value;
            }
        }
    }

    /// <summary>
    ///     Gets the available user settings.
    /// </summary>
    /// <value>The available user settings.</value>
    internal static IEnumerable<UserSettingName> AvailableUserSettings => ConverterDictionary.Keys;
}
