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
using System.Security.Cryptography.Xml;
using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     A wrapper class to facilitate creating XML signatures around wsu:Id.
/// </summary>
internal class WsSecurityUtilityIdSignedXml : SignedXml
{
    private static long _nextId;
    private static readonly string CommonPrefix = "uuid-" + Guid.NewGuid() + "-";

    private readonly XmlDocument _document;
    private readonly Dictionary<string, XmlElement> _ids;

    /// <summary>
    ///     Initializes a new instance of the WSSecurityUtilityIdSignedXml class from the specified XML document.
    /// </summary>
    /// <param name="document">Xml document.</param>
    public WsSecurityUtilityIdSignedXml(XmlDocument document)
        : base(document)
    {
        _document = document;
        _ids = new Dictionary<string, XmlElement>();
    }

    /// <summary>
    ///     Get unique Id.
    /// </summary>
    /// <returns>The wsu id.</returns>
    public static string GetUniqueId()
    {
        return CommonPrefix + Interlocked.Increment(ref _nextId).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    ///     Add the node as reference.
    /// </summary>
    /// <param name="xpath">The XPath string.</param>
    public void AddReference(string xpath)
    {
        // for now, ignore the error if the node is not found. 
        // EWS may want to sign extra header while such header is never present in autodiscover request.
        // but currently Credentials are unaware of the service type.
        // 
        if (_document.SelectSingleNode(xpath, WSSecurityBasedCredentials.NamespaceManager) is XmlElement element)
        {
            var wsuId = GetUniqueId();

            var wsuIdAttribute = _document.CreateAttribute(
                EwsUtilities.WsSecurityUtilityNamespacePrefix,
                "Id",
                EwsUtilities.WsSecurityUtilityNamespace
            );

            wsuIdAttribute.Value = wsuId;
            element.Attributes.Append(wsuIdAttribute);

            var reference = new Reference
            {
                Uri = "#" + wsuId,
            };
            reference.AddTransform(new XmlDsigExcC14NTransform());

            AddReference(reference);
            _ids.Add(wsuId, element);
        }
    }

    /// <summary>
    ///     Returns the XmlElement  object with the specified ID from the specified XmlDocument  object.
    /// </summary>
    /// <param name="document">The XmlDocument object to retrieve the XmlElement object from</param>
    /// <param name="idValue">The ID of the XmlElement object to retrieve from the XmlDocument object.</param>
    /// <returns>The XmlElement object with the specified ID from the specified XmlDocument object</returns>
    public override XmlElement GetIdElement(XmlDocument document, string idValue)
    {
        return _ids[idValue];
    }
}
