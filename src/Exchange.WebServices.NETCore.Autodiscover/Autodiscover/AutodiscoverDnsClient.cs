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

using System.Net;

using DnsClient;
using DnsClient.Protocol;

using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Dns;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Class that reads AutoDiscover configuration information from DNS.
/// </summary>
internal class AutodiscoverDnsClient
{
    /// <summary>
    ///     SRV DNS prefix to lookup.
    /// </summary>
    private const string AutoDiscoverSrvPrefix = "_autodiscover._tcp.";

    /// <summary>
    ///     We are only interested in records that use SSL.
    /// </summary>
    private const int SslPort = 443;


    /// <summary>
    ///     Random selector in the case of ties.
    /// </summary>
    private static readonly Random RandomTieBreakerSelector = new();


    /// <summary>
    ///     AutodiscoverService using this DNS reader.
    /// </summary>
    private readonly AutodiscoverService _service;


    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverDnsClient" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal AutodiscoverDnsClient(AutodiscoverService service)
    {
        _service = service;
    }


    /// <summary>
    ///     Finds the Autodiscover host from DNS SRV records.
    /// </summary>
    /// <remarks>
    ///     If the domain to lookup is "contoso.com", Autodiscover will use DnsQuery on SRV records
    ///     for "_autodiscover._tcp.contoso.com". If the query is successful it will return a target
    ///     domain (e.g. "mail.contoso.com") which will be tried as an Autodiscover endpoint.
    /// </remarks>
    /// <param name="domain">The domain.</param>
    /// <returns>Autodiscover hostname (will be null if lookup failed).</returns>
    internal async Task<string?> FindAutodiscoverHostFromSrv(string domain)
    {
        var domainToMatch = AutoDiscoverSrvPrefix + domain;

        var dnsSrvRecord = await FindBestMatchingSrvRecord(domainToMatch);

        if (dnsSrvRecord == null || string.IsNullOrEmpty(dnsSrvRecord.Target.Value))
        {
            _service.TraceMessage(TraceFlags.AutodiscoverConfiguration, "No appropriate SRV record was found.");
            return null;
        }

        _service.TraceMessage(
            TraceFlags.AutodiscoverConfiguration,
            $"DNS query for SRV record for domain {domain} found {dnsSrvRecord.Target.Value}"
        );

        return dnsSrvRecord.Target.Value;
    }

    /// <summary>
    ///     Finds the best matching SRV record.
    /// </summary>
    /// <param name="domain">The domain.</param>
    /// <returns>DnsSrvRecord(will be null if lookup failed).</returns>
    private async Task<SrvRecord?> FindBestMatchingSrvRecord(string domain)
    {
        // Make DnsQuery call to get collection of SRV records.
        var dnsSrvRecordList = await DnsQuery(domain, _service.DnsServerAddress);


        _service.TraceMessage(
            TraceFlags.AutodiscoverConfiguration,
            $"{dnsSrvRecordList.Count} SRV records were returned."
        );

        // If multiple records were returned, they will be returned sorted by priority 
        // (and weight) order. Need to find the index of the first record that supports SSL.
        var priority = int.MinValue;
        var weight = int.MinValue;
        var recordFound = false;
        foreach (var dnsSrvRecord in dnsSrvRecordList)
        {
            if (dnsSrvRecord.Port == SslPort)
            {
                priority = dnsSrvRecord.Priority;
                weight = dnsSrvRecord.Weight;
                recordFound = true;
                break;
            }
        }

        // Records were returned but nothing matched our criteria.
        if (!recordFound)
        {
            _service.TraceMessage(TraceFlags.AutodiscoverConfiguration, "No appropriate SRV records were found.");
            return null;
        }

        // Collect all records with the same (highest) priority.
        // (Aren't lambda expressions cool? ;-)
        var bestDnsSrvRecordList = dnsSrvRecordList.FindAll(
            record => record.Port == SslPort && record.Priority == priority && record.Weight == weight
        );

        // The list must contain at least one matching record since we found one earlier.
        EwsUtilities.Assert(
            dnsSrvRecordList.Count > 0,
            "AutodiscoverDnsClient.FindBestMatchingSrvRecord",
            "At least one DNS SRV record must match the criteria."
        );

        // If we have multiple records with the same priority and weight, randomly pick one.
        var recordIndex = bestDnsSrvRecordList.Count > 1 ? RandomTieBreakerSelector.Next(bestDnsSrvRecordList.Count)
            : 0;

        var bestDnsSrvRecord = bestDnsSrvRecordList[recordIndex];

        var traceMessage = string.Format(
            "Returning SRV record {0} of {1} records. Target: {2}, Priority: {3}, Weight: {4}",
            recordIndex,
            dnsSrvRecordList.Count,
            bestDnsSrvRecord.Target,
            bestDnsSrvRecord.Priority,
            bestDnsSrvRecord.Weight
        );
        _service.TraceMessage(TraceFlags.AutodiscoverConfiguration, traceMessage);

        return bestDnsSrvRecord;
    }

    private static async Task<List<SrvRecord>> DnsQuery(string domain, IPAddress? dnsServerAddress)
    {
        var options = dnsServerAddress != null ? new LookupClientOptions(dnsServerAddress) : new LookupClientOptions();

        var lookup = new LookupClient(options);

        var response = await lookup.QueryAsync(domain, QueryType.SRV);

        if (response.HasError)
        {
            return new List<SrvRecord>();
        }

        return response.Answers.SrvRecords().ToList();
    }
}
