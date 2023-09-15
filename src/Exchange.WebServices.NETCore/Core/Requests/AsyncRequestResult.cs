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
///     IAsyncResult implementation to be returned to caller - decorator pattern.
/// </summary>
internal class AsyncRequestResult : IAsyncResult
{
    /// <summary>
    ///     Constructor
    /// </summary>
    /// <param name="serviceRequest"></param>
    /// <param name="webRequest"></param>
    /// <param name="webAsyncResult"></param>
    /// <param name="asyncState"></param>
    public AsyncRequestResult(
        ServiceRequestBase serviceRequest,
        IEwsHttpWebRequest webRequest,
        IAsyncResult webAsyncResult,
        object asyncState
    )
    {
        EwsUtilities.ValidateParam(serviceRequest);
        EwsUtilities.ValidateParam(webRequest);
        EwsUtilities.ValidateParam(webAsyncResult);

        ServiceRequest = serviceRequest;
        WebAsyncResult = webAsyncResult;
        WebRequest = webRequest;
        AsyncState = asyncState;
    }

    /// <summary>
    ///     ServiceRequest
    /// </summary>
    public ServiceRequestBase ServiceRequest { get; private set; }

    /// <summary>
    ///     WebRequest
    /// </summary>
    public IEwsHttpWebRequest WebRequest { get; private set; }

    /// <summary>
    ///     AsyncResult
    /// </summary>
    public IAsyncResult WebAsyncResult { get; private set; }

    /// <summary>
    ///     AsyncState
    /// </summary>
    public object AsyncState { get; private set; }

    /// <summary>
    ///     AsyncWaitHandle
    /// </summary>
    public WaitHandle AsyncWaitHandle => WebAsyncResult.AsyncWaitHandle;

    /// <summary>
    ///     CompletedSynchronously
    /// </summary>
    public bool CompletedSynchronously => WebAsyncResult.CompletedSynchronously;

    /// <summary>
    ///     IsCompleted
    /// </summary>
    public bool IsCompleted => WebAsyncResult.IsCompleted;

    /// <summary>
    ///     Extracts the original service request from the specified IAsyncResult instance
    /// </summary>
    /// <typeparam name="T">Desired service request type</typeparam>
    /// <param name="exchangeService">The ExchangeService object to validate the integrity of asyncResult</param>
    /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
    /// <returns>The original service request</returns>
    public static T ExtractServiceRequest<T>(ExchangeService exchangeService, IAsyncResult asyncResult)
        where T : SimpleServiceRequestBase
    {
        // Validate null first
        EwsUtilities.ValidateParam(asyncResult);

        if (asyncResult is not AsyncRequestResult asyncRequestResult)
        {
            // Strings.InvalidAsyncResult is copied from the error message of HttpWebRequest.EndGetResponse()
            // Just use this simple string for all kinds of invalid IAsyncResult parameters
            throw new ArgumentException(Strings.InvalidAsyncResult, nameof(asyncResult));
        }

        // Validate the service request
        if (asyncRequestResult.ServiceRequest == null)
        {
            throw new ArgumentException(Strings.InvalidAsyncResult, nameof(asyncResult));
        }

        //Validate the service object
        if (!ReferenceEquals(asyncRequestResult.ServiceRequest.Service, exchangeService))
        {
            throw new ArgumentException(Strings.InvalidAsyncResult, nameof(asyncResult));
        }

        // Validate the request type
        if (asyncRequestResult.ServiceRequest is not T serviceRequest)
        {
            throw new ArgumentException(Strings.InvalidAsyncResult, nameof(asyncResult));
        }

        return serviceRequest;
    }
}

/// <summary>
///     State object wrapper to be passed to HttpWebRequest's async methods
/// </summary>
internal class WebAsyncCallStateAnchor
{
    /// <summary>
    ///     Constructor
    /// </summary>
    /// <param name="serviceRequest"></param>
    /// <param name="webRequest"></param>
    /// <param name="asyncCallback"></param>
    /// <param name="asyncState"></param>
    public WebAsyncCallStateAnchor(
        ServiceRequestBase serviceRequest,
        IEwsHttpWebRequest webRequest,
        AsyncCallback asyncCallback,
        object asyncState
    )
    {
        EwsUtilities.ValidateParam(serviceRequest);
        EwsUtilities.ValidateParam(webRequest);

        ServiceRequest = serviceRequest;
        WebRequest = webRequest;

        AsyncCallback = asyncCallback;
        AsyncState = asyncState;
    }

    /// <summary>
    ///     ServiceRequest
    /// </summary>
    public ServiceRequestBase ServiceRequest { get; private set; }

    /// <summary>
    ///     WebRequest
    /// </summary>
    public IEwsHttpWebRequest WebRequest { get; private set; }

    /// <summary>
    ///     AsyncState
    /// </summary>
    public object AsyncState { get; private set; }

    /// <summary>
    ///     AsyncCallback
    /// </summary>
    public AsyncCallback AsyncCallback { get; private set; }
}