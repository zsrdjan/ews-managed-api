using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
/// Well known mapi properties which are defined in the MS-OXPROPS list.
/// </summary>
/// <remarks>
/// See more: https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/
/// </remarks>
[PublicAPI]
public static class KnownMapiProperties
{
    /// <summary>
    /// Specifies the hide or show status of a folder.
    /// </summary>
    /// <remarks></remarks>
    public static readonly ExtendedPropertyDefinition AttributeHidden = new(4340, MapiPropertyType.Boolean);
}
