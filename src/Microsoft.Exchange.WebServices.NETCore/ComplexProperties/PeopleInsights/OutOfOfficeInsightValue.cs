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

using System.Globalization;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the OutOfOfficeInsightValue.
/// </summary>
[PublicAPI]
public sealed class OutOfOfficeInsightValue : InsightValue
{
    /// <summary>
    ///     Get the start date and time.
    /// </summary>
    public DateTime StartTime { get; private set; }

    /// <summary>
    ///     Get the end date and time.
    /// </summary>
    public DateTime EndTime { get; private set; }

    /// <summary>
    ///     Get the culture of the reply.
    /// </summary>
    public string Culture { get; private set; } = CultureInfo.CurrentCulture.Name;

    /// <summary>
    ///     Get the reply message.
    /// </summary>
    public string Message { get; private set; }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">XML reader</param>
    /// <returns>Whether the element was read</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.InsightSource:
            {
                InsightSource = reader.ReadElementValue<string>();
                break;
            }
            case XmlElementNames.StartTime:
            {
                StartTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.StartTime).Value;
                break;
            }
            case XmlElementNames.EndTime:
            {
                EndTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.EndTime).Value;
                break;
            }
            case XmlElementNames.Culture:
            {
                Culture = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Message:
            {
                Message = reader.ReadElementValue();
                break;
            }
            default:
            {
                return false;
            }
        }

        return true;
    }
}
