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
///     Represents the EducationInsightValue.
/// </summary>
[PublicAPI]
public sealed class EducationInsightValue : InsightValue
{
    private string _degree;
    private long _endUtcTicks;
    private string _institute;
    private long _startUtcTicks;

    /// <summary>
    ///     Gets the Institute
    /// </summary>
    public string Institute
    {
        get => _institute;
        set => SetFieldValue(ref _institute, value);
    }

    /// <summary>
    ///     Gets the Degree
    /// </summary>
    public string Degree
    {
        get => _degree;
        set => SetFieldValue(ref _degree, value);
    }

    /// <summary>
    ///     Gets the StartUtcTicks
    /// </summary>
    public long StartUtcTicks
    {
        get => _startUtcTicks;
        set => SetFieldValue(ref _startUtcTicks, value);
    }

    /// <summary>
    ///     Gets the EndUtcTicks
    /// </summary>
    public long EndUtcTicks
    {
        get => _endUtcTicks;
        set => SetFieldValue(ref _endUtcTicks, value);
    }

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
            case XmlElementNames.UpdatedUtcTicks:
            {
                UpdatedUtcTicks = reader.ReadElementValue<long>();
                break;
            }
            case XmlElementNames.Institute:
            {
                Institute = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Degree:
            {
                Degree = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.StartUtcTicks:
            {
                StartUtcTicks = reader.ReadElementValue<long>();
                break;
            }
            case XmlElementNames.EndUtcTicks:
            {
                EndUtcTicks = reader.ReadElementValue<long>();
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
