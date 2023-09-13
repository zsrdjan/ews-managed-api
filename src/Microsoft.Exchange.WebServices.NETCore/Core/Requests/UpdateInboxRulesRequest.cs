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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a UpdateInboxRulesRequest request.
/// </summary>
internal sealed class UpdateInboxRulesRequest : SimpleServiceRequestBase
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="UpdateInboxRulesRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal UpdateInboxRulesRequest(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.UpdateInboxRules;
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (!string.IsNullOrEmpty(MailboxSmtpAddress))
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MailboxSmtpAddress, MailboxSmtpAddress);
        }

        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.RemoveOutlookRuleBlob, RemoveOutlookRuleBlob);
        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Operations);
        foreach (var operation in InboxRuleOperations)
        {
            operation.WriteToXml(writer, operation.XmlElementName);
        }

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.UpdateInboxRulesResponse;
    }

    /// <summary>
    ///     Parses the response.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Response object.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader)
    {
        var response = new UpdateInboxRulesResponse();
        response.LoadFromXml(reader, XmlElementNames.UpdateInboxRulesResponse);
        return response;
    }

    /// <summary>
    ///     Gets the request version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this request is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2010_SP1;
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        if (InboxRuleOperations == null)
        {
            throw new ArgumentException("RuleOperations cannot be null.", "Operations");
        }

        var operationCount = 0;
        foreach (var operation in InboxRuleOperations)
        {
            EwsUtilities.ValidateParam(operation, "RuleOperation");
            operationCount++;
        }

        if (operationCount == 0)
        {
            throw new ArgumentException("RuleOperations cannot be empty.", "Operations");
        }

        Service.Validate();
    }

    /// <summary>
    ///     Executes this request.
    /// </summary>
    /// <returns>Service response.</returns>
    internal async Task<UpdateInboxRulesResponse> Execute(CancellationToken token)
    {
        var serviceResponse = (UpdateInboxRulesResponse)await InternalExecuteAsync(token).ConfigureAwait(false);
        if (serviceResponse.Result == ServiceResult.Error)
        {
            throw new UpdateInboxRulesException(serviceResponse, InboxRuleOperations.GetEnumerator());
        }

        return serviceResponse;
    }

    /// <summary>
    ///     Gets or sets the address of the mailbox in which to update the inbox rules.
    /// </summary>
    internal string MailboxSmtpAddress { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether or not to remove OutlookRuleBlob from
    ///     the rule collection.
    /// </summary>
    internal bool RemoveOutlookRuleBlob { get; set; }

    /// <summary>
    ///     Gets or sets the RuleOperation collection.
    /// </summary>
    internal IEnumerable<RuleOperation> InboxRuleOperations { get; set; }
}
