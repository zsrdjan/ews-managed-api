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
///     Represents the JobInsightValue.
/// </summary>
[PublicAPI]
public sealed class JobInsightValue : InsightValue
{
    private string _company;
    private string _companyDescription;
    private string _companyTicker;
    private string _companyLogoUrl;
    private string _companyWebsiteUrl;
    private string _companyLinkedInUrl;
    private string _title;
    private long _startUtcTicks;
    private long _endUtcTicks;

    /// <summary>
    ///     Gets the Company
    /// </summary>
    public string Company
    {
        get => _company;
        set => SetFieldValue(ref _company, value);
    }

    /// <summary>
    ///     Gets the CompanyDescription
    /// </summary>
    public string CompanyDescription
    {
        get => _companyDescription;
        set => SetFieldValue(ref _companyDescription, value);
    }

    /// <summary>
    ///     Gets the CompanyTicker
    /// </summary>
    public string CompanyTicker
    {
        get => _companyTicker;
        set => SetFieldValue(ref _companyTicker, value);
    }

    /// <summary>
    ///     Gets the CompanyLogoUrl
    /// </summary>
    public string CompanyLogoUrl
    {
        get => _companyLogoUrl;
        set => SetFieldValue(ref _companyLogoUrl, value);
    }

    /// <summary>
    ///     Gets the CompanyWebsiteUrl
    /// </summary>
    public string CompanyWebsiteUrl
    {
        get => _companyWebsiteUrl;
        set => SetFieldValue(ref _companyWebsiteUrl, value);
    }

    /// <summary>
    ///     Gets the CompanyLinkedInUrl
    /// </summary>
    public string CompanyLinkedInUrl
    {
        get => _companyLinkedInUrl;
        set => SetFieldValue(ref _companyLinkedInUrl, value);
    }

    /// <summary>
    ///     Gets the Title
    /// </summary>
    public string Title
    {
        get => _title;
        set => SetFieldValue(ref _title, value);
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
            case XmlElementNames.Company:
            {
                Company = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Title:
            {
                Title = reader.ReadElementValue();
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
