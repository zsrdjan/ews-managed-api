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

using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents SoapFault details.
/// </summary>
internal class SoapFaultDetails
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="SoapFaultDetails" /> class.
    /// </summary>
    private SoapFaultDetails()
    {
    }


    /// <summary>
    ///     Parses the soap:Fault content.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="soapNamespace">The SOAP namespace to use.</param>
    /// <returns>SOAP fault details.</returns>
    internal static SoapFaultDetails Parse(EwsXmlReader reader, XmlNamespace soapNamespace)
    {
        var soapFaultDetails = new SoapFaultDetails();

        do
        {
            reader.Read();
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.SOAPFaultCodeElementName:
                    {
                        soapFaultDetails.FaultCode = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.SOAPFaultStringElementName:
                    {
                        soapFaultDetails.FaultString = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.SOAPFaultActorElementName:
                    {
                        soapFaultDetails.FaultActor = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.SOAPDetailElementName:
                    {
                        soapFaultDetails.ParseDetailNode(reader);
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPFaultElementName));

        return soapFaultDetails;
    }

    /// <summary>
    ///     Parses the detail node.
    /// </summary>
    /// <param name="reader">The reader.</param>
    private void ParseDetailNode(EwsXmlReader reader)
    {
        do
        {
            reader.Read();
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.EwsResponseCodeElementName:
                    {
                        try
                        {
                            ResponseCode = reader.ReadElementValue<ServiceError>();
                        }
                        catch (ArgumentException)
                        {
                            // ServiceError couldn't be mapped to enum value, treat as an ISE
                            ResponseCode = ServiceError.ErrorInternalServerError;
                        }

                        break;
                    }
                    case XmlElementNames.EwsMessageElementName:
                    {
                        Message = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EwsLineElementName:
                    {
                        LineNumber = reader.ReadElementValue<int>();
                        break;
                    }
                    case XmlElementNames.EwsPositionElementName:
                    {
                        PositionWithinLine = reader.ReadElementValue<int>();
                        break;
                    }
                    case XmlElementNames.EwsErrorCodeElementName:
                    {
                        try
                        {
                            ErrorCode = reader.ReadElementValue<ServiceError>();
                        }
                        catch (ArgumentException)
                        {
                            // ServiceError couldn't be mapped to enum value, treat as an ISE
                            ErrorCode = ServiceError.ErrorInternalServerError;
                        }

                        break;
                    }
                    case XmlElementNames.EwsExceptionTypeElementName:
                    {
                        ExceptionType = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.MessageXml:
                    {
                        ParseMessageXml(reader);
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.SOAPDetailElementName));
    }

    /// <summary>
    ///     Parses the message XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    private void ParseMessageXml(EwsXmlReader reader)
    {
        // E12 and E14 return the MessageXml element in different
        // namespaces (types namespace for E12, errors namespace in E14). To
        // avoid this problem, the parser will match the namespace from the
        // start and end elements.
        var elementNs = EwsUtilities.GetNamespaceFromUri(reader.NamespaceUri);

        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.IsStartElement() && !reader.IsEmptyElement)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.Value:
                        {
                            ErrorDetails.Add(
                                reader.ReadAttributeValue(XmlAttributeNames.Name),
                                reader.ReadElementValue()
                            );
                            break;
                        }
                    }
                }
            } while (!reader.IsEndElement(elementNs, XmlElementNames.MessageXml));
        }
    }

    /// <summary>
    ///     Gets or sets the SOAP fault code.
    /// </summary>
    /// <value>The SOAP fault code.</value>
    internal string FaultCode { get; set; }

    /// <summary>
    ///     Gets or sets the SOAP fault string.
    /// </summary>
    /// <value>The fault string.</value>
    internal string FaultString { get; set; }

    /// <summary>
    ///     Gets or sets the SOAP fault actor.
    /// </summary>
    /// <value>The fault actor.</value>
    internal string FaultActor { get; set; }

    /// <summary>
    ///     Gets or sets the response code returned by EWS requests.
    /// </summary>
    /// <remarks>Default to InternalServerError.</remarks>
    /// <value>The response code.</value>
    internal ServiceError ResponseCode { get; set; } = ServiceError.ErrorInternalServerError;

    /// <summary>
    ///     Gets or sets the message.
    /// </summary>
    /// <value>The message.</value>
    internal string Message { get; set; }

    /// <summary>
    ///     Gets or sets the error code.
    /// </summary>
    /// <value>The error code.</value>
    internal ServiceError ErrorCode { get; set; } = ServiceError.NoError;

    /// <summary>
    ///     Gets or sets the type of the exception.
    /// </summary>
    /// <remarks>This is returned by UM requests. It's the name of the exception that was raised.</remarks>
    /// <value>The type of the exception.</value>
    internal string ExceptionType { get; set; }

    /// <summary>
    ///     Gets or sets the line number.
    /// </summary>
    /// <remarks>When a schema validation error is returned, this is the line number in the request where the error occurred.</remarks>
    /// <value>The line number.</value>
    internal int LineNumber { get; set; }

    /// <summary>
    ///     Gets or sets the position within line.
    /// </summary>
    /// <remarks>
    /// When a schema validation error is returned, this is the offset
    /// into the line of the request where the error occurred.
    /// </remarks>
    /// <value>The position within line.</value>
    internal int PositionWithinLine { get; set; }

    /// <summary>
    ///     Gets or sets the error details dictionary.
    /// </summary>
    /// <remarks>
    ///     Dictionary of key/value pairs from the MessageXml node in the fault. Usually empty but there are
    ///     a few cases where SOAP faults may include MessageXml details (e.g. CASOverBudgetException includes
    ///     BackoffTime value).
    /// </remarks>
    /// <value>The error details dictionary.</value>
    internal Dictionary<string, string> ErrorDetails { get; set; } = new();
}
