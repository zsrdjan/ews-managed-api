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

using System.Net.Http.Headers;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an abstract, simple request-response service request.
/// </summary>
internal abstract class SimpleServiceRequestBase : ServiceRequestBase
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="SimpleServiceRequestBase" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal SimpleServiceRequestBase(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Executes this request.
    /// </summary>
    /// <returns>Service response.</returns>
    internal async Task<object> InternalExecuteAsync(CancellationToken token)
    {
        var tuple = await ValidateAndEmitRequest(token).ConfigureAwait(false);
        try
        {
            return await ReadResponse(tuple.Item2);
        }
        finally
        {
            tuple.Item1.Dispose();
            tuple.Item2.Dispose();
        }
    }

    /// <summary>
    ///     Async callback method for HttpWebRequest async requests.
    /// </summary>
    /// <param name="webAsyncResult">An IAsyncResult that references the asynchronous request.</param>
    private static void WebRequestAsyncCallback(IAsyncResult webAsyncResult)
    {
        var wrappedState = webAsyncResult.AsyncState as WebAsyncCallStateAnchor;

        if (wrappedState != null && wrappedState.AsyncCallback != null)
        {
            var asyncRequestResult = new AsyncRequestResult(
                wrappedState.ServiceRequest,
                wrappedState.WebRequest,
                webAsyncResult, /* web async result */
                wrappedState.AsyncState /* user state */
            );

            // Call user's call back
            wrappedState.AsyncCallback(asyncRequestResult);
        }
    }

    /// <summary>
    ///     Reads the response with error handling
    /// </summary>
    /// <param name="response">The response.</param>
    /// <returns>Service response.</returns>
    private async Task<object> ReadResponse(IEwsHttpWebResponse response)
    {
        object serviceResponse;

        try
        {
            Service.ProcessHttpResponseHeaders(TraceFlags.EwsResponseHttpHeaders, response);

            // If tracing is enabled, we read the entire response into a MemoryStream so that we
            // can pass it along to the ITraceListener. Then we parse the response from the 
            // MemoryStream.
            if (Service.IsTraceEnabledFor(TraceFlags.EwsResponse))
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var serviceResponseStream = await GetResponseStream(response))
                    {
                        // Copy response to in-memory stream and reset position to start.
                        EwsUtilities.CopyStream(serviceResponseStream, memoryStream);
                        memoryStream.Position = 0;
                    }

                    TraceResponseXml(response, memoryStream);

                    serviceResponse = ReadResponseXml(memoryStream, response.Headers);
                }
            }
            else
            {
                using (var responseStream = await GetResponseStream(response))
                {
                    serviceResponse = ReadResponseXml(responseStream, response.Headers);
                }
            }
        }
        catch (EwsHttpClientException e)
        {
            if (e.Response != null)
            {
                var exceptionResponse = Service.HttpWebRequestFactory.CreateExceptionResponse(e);
                Service.ProcessHttpResponseHeaders(TraceFlags.EwsResponseHttpHeaders, exceptionResponse);
            }

            throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
        }
        catch (IOException e)
        {
            // Wrap exception.
            throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
        }
        finally
        {
            if (response != null)
            {
                response.Close();
            }
        }

        return serviceResponse;
    }

    /// <summary>
    ///     Reads the response XML.
    /// </summary>
    /// <param name="responseStream">The response stream.</param>
    /// <returns></returns>
    private object ReadResponseXml(Stream responseStream)
    {
        return ReadResponseXml(responseStream, null);
    }

    /// <summary>
    ///     Reads the response XML.
    /// </summary>
    /// <param name="responseStream">The response stream.</param>
    /// <param name="responseHeaders">The HTTP response headers</param>
    /// <returns></returns>
    private object ReadResponseXml(Stream responseStream, HttpResponseHeaders responseHeaders)
    {
        object serviceResponse;
        var ewsXmlReader = new EwsServiceXmlReader(responseStream, Service);
        serviceResponse = ReadResponse(ewsXmlReader, responseHeaders);
        return serviceResponse;
    }
}
