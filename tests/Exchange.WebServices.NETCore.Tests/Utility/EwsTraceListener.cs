using Microsoft.Exchange.WebServices.Data;

namespace Exchange.WebServices.NETCore.Tests.Utility;

internal class EwsTraceListener : ITraceListener
{
    public void Trace(string traceType, string traceMessage)
    {
        Console.WriteLine("{0} {1}", traceType, traceMessage);
    }
}
