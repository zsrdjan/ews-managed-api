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
///     Represents a phone call.
/// </summary>
public sealed class PhoneCall : ComplexProperty
{
    private const string SuccessfulResponseText = "OK";
    private const int SuccessfulResponseCode = 200;

    private readonly ExchangeService _service;
    private readonly PhoneCallId _id;

    /// <summary>
    ///     PhoneCall Constructor.
    /// </summary>
    /// <param name="service">EWS service to which this object belongs.</param>
    internal PhoneCall(ExchangeService service)
    {
        EwsUtilities.Assert(service != null, "PhoneCall.ctor", "service is null");

        _service = service;
        State = PhoneCallState.Connecting;
        ConnectionFailureCause = ConnectionFailureCause.None;
        SIPResponseText = SuccessfulResponseText;
        SIPResponseCode = SuccessfulResponseCode;
    }

    /// <summary>
    ///     PhoneCall Constructor.
    /// </summary>
    /// <param name="service">EWS service to which this object belongs.</param>
    /// <param name="id">The Id of the phone call.</param>
    internal PhoneCall(ExchangeService service, PhoneCallId id)
        : this(service)
    {
        _id = id;
    }

    /// <summary>
    ///     Refreshes the state of this phone call.
    /// </summary>
    public async System.Threading.Tasks.Task Refresh(CancellationToken token = default)
    {
        var phoneCall = await _service.UnifiedMessaging.GetPhoneCallInformation(_id, token).ConfigureAwait(false);
        State = phoneCall.State;
        ConnectionFailureCause = phoneCall.ConnectionFailureCause;
        SIPResponseText = phoneCall.SIPResponseText;
        SIPResponseCode = phoneCall.SIPResponseCode;
    }

    /// <summary>
    ///     Disconnects this phone call.
    /// </summary>
    public async System.Threading.Tasks.Task Disconnect(CancellationToken token = default)
    {
        // If call is already disconnected, throw exception
        //
        if (State == PhoneCallState.Disconnected)
        {
            throw new ServiceLocalException(Strings.PhoneCallAlreadyDisconnected);
        }

        await _service.UnifiedMessaging.DisconnectPhoneCall(_id, token);
        State = PhoneCallState.Disconnected;
    }

    /// <summary>
    ///     Tries to read an element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.PhoneCallState:
                State = reader.ReadElementValue<PhoneCallState>();
                return true;
            case XmlElementNames.ConnectionFailureCause:
                ConnectionFailureCause = reader.ReadElementValue<ConnectionFailureCause>();
                return true;
            case XmlElementNames.SIPResponseText:
                SIPResponseText = reader.ReadElementValue();
                return true;
            case XmlElementNames.SIPResponseCode:
                SIPResponseCode = reader.ReadElementValue<int>();
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    ///     Gets a value indicating the last known state of this phone call.
    /// </summary>
    public PhoneCallState State { get; private set; }

    /// <summary>
    ///     Gets a value indicating the reason why this phone call failed to connect.
    /// </summary>
    public ConnectionFailureCause ConnectionFailureCause { get; private set; }

    /// <summary>
    ///     Gets the SIP response text of this phone call.
    /// </summary>
    public string SIPResponseText { get; private set; }

    /// <summary>
    ///     Gets the SIP response code of this phone call.
    /// </summary>
    public int SIPResponseCode { get; private set; }
}
