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
///     Represents an abstract Move/Copy Folder request.
/// </summary>
/// <typeparam name="TResponse">The type of the response.</typeparam>
internal abstract class MoveCopyFolderRequest<TResponse> : MoveCopyRequest<Folder, TResponse>
    where TResponse : ServiceResponse
{
    /// <summary>
    ///     Gets the folder ids.
    /// </summary>
    /// <value>The folder ids.</value>
    internal FolderIdWrapperList FolderIds { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="MoveCopyFolderRequest&lt;TResponse&gt;" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    internal MoveCopyFolderRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service, errorHandlingMode)
    {
    }

    /// <summary>
    ///     Validates request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();
        EwsUtilities.ValidateParamCollection(FolderIds);
        FolderIds.Validate(Service.RequestedServerVersion);
    }

    /// <summary>
    ///     Writes the ids as XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteIdsToXml(EwsServiceXmlWriter writer)
    {
        FolderIds.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.FolderIds);
    }

    /// <summary>
    ///     Gets the expected response message count.
    /// </summary>
    /// <returns>Number of expected response messages.</returns>
    protected override int GetExpectedResponseMessageCount()
    {
        return FolderIds.Count;
    }
}
