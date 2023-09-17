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
using System.Runtime.InteropServices;

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Dns;

/// <summary>
///     DNS Query client.
/// </summary>
internal class DnsClient
{
    /// <summary>
    ///     Win32 successful operation.
    /// </summary>
    private const int Win32Success = 0;

    /// <summary>
    ///     Map type of DnsRecord to DnsRecordType.
    /// </summary>
    private static readonly LazyMember<Dictionary<Type, DnsRecordType>> TypeToDnsTypeMap = new(
        () => new Dictionary<Type, DnsRecordType>
        {
            {
                typeof(DnsSrvRecord), DnsRecordType.SRV
            },
        }
    );

    /// <summary>
    ///     Perform DNS Query.
    /// </summary>
    /// <typeparam name="T">DnsRecord type.</typeparam>
    /// <param name="domain">The domain.</param>
    /// <param name="dnsServerAddress">IPAddress of DNS server to use (may be null).</param>
    /// <returns>The DNS record list (never null but may be empty).</returns>
    internal static List<T> DnsQuery<T>(string domain, IPAddress dnsServerAddress)
        where T : DnsRecord, new()
    {
        var dnsRecordList = new List<T>();

        // Each strongly-typed DnsRecord type maps to a DnsRecordType enum.
        var dnsRecordTypeToQuery = TypeToDnsTypeMap.Member[typeof(T)];

        // queryResultsPtr will point to unmanaged heap memory if DnsQuery succeeds.
        var queryResultsPtr = IntPtr.Zero;

        try
        {
            // Perform DNS query. If successful, construct a list of results.
            var errorCode = DnsNativeMethods.DnsQuery(
                domain,
                dnsServerAddress,
                dnsRecordTypeToQuery,
                ref queryResultsPtr
            );

            if (errorCode == Win32Success)
            {
                DnsRecordHeader dnsRecordHeader;

                // Iterate through linked list of query result records.
                for (var recordPtr = queryResultsPtr;
                     !recordPtr.Equals(IntPtr.Zero);
                     recordPtr = dnsRecordHeader.NextRecord)
                {
                    dnsRecordHeader = Marshal.PtrToStructure<DnsRecordHeader>(recordPtr);

                    var dnsRecord = new T();
                    if (dnsRecordHeader.RecordType == dnsRecord.RecordType)
                    {
                        dnsRecord.Load(dnsRecordHeader, recordPtr);
                        dnsRecordList.Add(dnsRecord);
                    }
                }
            }
            else
            {
                throw new DnsException(errorCode);
            }
        }
        finally
        {
            if (queryResultsPtr != IntPtr.Zero)
            {
                // DnsQuery allocated unmanaged heap, free it now.
                DnsNativeMethods.FreeDnsQueryResults(queryResultsPtr);
            }
        }

        return dnsRecordList;
    }
}
