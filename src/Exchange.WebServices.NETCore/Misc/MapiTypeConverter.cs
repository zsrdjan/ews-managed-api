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

using System.Globalization;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Utility class to convert between MAPI Property type values and strings.
/// </summary>
internal static class MapiTypeConverter
{
    /// <summary>
    ///     Assume DateTime values are in UTC.
    /// </summary>
    private const DateTimeStyles UtcDataTimeStyles = DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal;

    /// <summary>
    ///     Map from MAPI property type to converter entry.
    /// </summary>
    internal static readonly IReadOnlyDictionary<MapiPropertyType, MapiTypeConverterMapEntry> MapiTypeConverterMap =
        new Dictionary<MapiPropertyType, MapiTypeConverterMapEntry>
        {
            {
                MapiPropertyType.ApplicationTime, new MapiTypeConverterMapEntry(typeof(double))
            },
            {
                MapiPropertyType.ApplicationTimeArray, new MapiTypeConverterMapEntry(typeof(double))
                {
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Binary, new MapiTypeConverterMapEntry(typeof(byte[]))
                {
                    Parse = s => string.IsNullOrEmpty(s) ? null : Convert.FromBase64String(s),
                    ConvertToString = o => Convert.ToBase64String((byte[])o),
                }
            },
            {
                MapiPropertyType.BinaryArray, new MapiTypeConverterMapEntry(typeof(byte[]))
                {
                    Parse = s => string.IsNullOrEmpty(s) ? null : Convert.FromBase64String(s),
                    ConvertToString = o => Convert.ToBase64String((byte[])o),
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Boolean, new MapiTypeConverterMapEntry(typeof(bool))
                {
                    Parse = s => Convert.ChangeType(s, typeof(bool), CultureInfo.InvariantCulture),
                    ConvertToString = o => ((bool)o).ToString().ToLower(),
                }
            },
            {
                MapiPropertyType.CLSID, new MapiTypeConverterMapEntry(typeof(Guid))
                {
                    Parse = s => new Guid(s),
                    ConvertToString = o => ((Guid)o).ToString(),
                }
            },
            {
                MapiPropertyType.CLSIDArray, new MapiTypeConverterMapEntry(typeof(Guid))
                {
                    Parse = s => new Guid(s),
                    ConvertToString = o => ((Guid)o).ToString(),
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Currency, new MapiTypeConverterMapEntry(typeof(long))
            },
            {
                MapiPropertyType.CurrencyArray, new MapiTypeConverterMapEntry(typeof(long))
                {
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Double, new MapiTypeConverterMapEntry(typeof(double))
            },
            {
                MapiPropertyType.DoubleArray, new MapiTypeConverterMapEntry(typeof(double))
                {
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Error, new MapiTypeConverterMapEntry(typeof(int))
            },
            {
                MapiPropertyType.Float, new MapiTypeConverterMapEntry(typeof(float))
            },
            {
                MapiPropertyType.FloatArray, new MapiTypeConverterMapEntry(typeof(float))
                {
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Integer, new MapiTypeConverterMapEntry(typeof(int))
                {
                    Parse = ParseMapiIntegerValue,
                }
            },
            {
                MapiPropertyType.IntegerArray, new MapiTypeConverterMapEntry(typeof(int))
                {
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Long, new MapiTypeConverterMapEntry(typeof(long))
            },
            {
                MapiPropertyType.LongArray, new MapiTypeConverterMapEntry(typeof(long))
                {
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Object, new MapiTypeConverterMapEntry(typeof(string))
                {
                    Parse = s => s,
                }
            },
            {
                MapiPropertyType.ObjectArray, new MapiTypeConverterMapEntry(typeof(string))
                {
                    Parse = s => s,
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.Short, new MapiTypeConverterMapEntry(typeof(short))
            },
            {
                MapiPropertyType.ShortArray, new MapiTypeConverterMapEntry(typeof(short))
                {
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.String, new MapiTypeConverterMapEntry(typeof(string))
                {
                    Parse = s => s,
                }
            },
            {
                MapiPropertyType.StringArray, new MapiTypeConverterMapEntry(typeof(string))
                {
                    Parse = s => s,
                    IsArray = true,
                }
            },
            {
                MapiPropertyType.SystemTime, new MapiTypeConverterMapEntry(typeof(DateTime))
                {
                    Parse = s => DateTime.Parse(s, CultureInfo.InvariantCulture, UtcDataTimeStyles),
                    ConvertToString =
                        o => EwsUtilities.DateTimeToXsDateTime((DateTime)o), // Can't use DataTime.ToString()
                }
            },
            {
                MapiPropertyType.SystemTimeArray, new MapiTypeConverterMapEntry(typeof(DateTime))
                {
                    IsArray = true,
                    Parse = s => DateTime.Parse(s, CultureInfo.InvariantCulture, UtcDataTimeStyles),
                    ConvertToString =
                        o => EwsUtilities.DateTimeToXsDateTime((DateTime)o), // Can't use DataTime.ToString()
                }
            },
        };


    /// <summary>
    ///     Converts the string list to array.
    /// </summary>
    /// <param name="mapiPropType">Type of the MAPI property.</param>
    /// <param name="strings">Strings.</param>
    /// <returns>Array of objects.</returns>
    internal static Array ConvertToValue(MapiPropertyType mapiPropType, IEnumerable<string> strings)
    {
        EwsUtilities.ValidateParam(strings);

        var typeConverter = MapiTypeConverterMap[mapiPropType];
        var array = Array.CreateInstance(typeConverter.Type, strings.Count());

        var index = 0;
        foreach (var stringValue in strings)
        {
            var value = typeConverter.ConvertToValueOrDefault(stringValue);
            array.SetValue(value, index++);
        }

        return array;
    }

    /// <summary>
    ///     Converts a string to value consistent with MAPI type.
    /// </summary>
    /// <param name="mapiPropType">Type of the MAPI property.</param>
    /// <param name="stringValue">String to convert to a value.</param>
    /// <returns></returns>
    internal static object ConvertToValue(MapiPropertyType mapiPropType, string stringValue)
    {
        return MapiTypeConverterMap[mapiPropType].ConvertToValue(stringValue);
    }

    /// <summary>
    ///     Converts a value to a string.
    /// </summary>
    /// <param name="mapiPropType">Type of the MAPI property.</param>
    /// <param name="value">Value to convert to string.</param>
    /// <returns>String value.</returns>
    internal static string ConvertToString(MapiPropertyType mapiPropType, object? value)
    {
        return value == null ? string.Empty : MapiTypeConverterMap[mapiPropType].ConvertToString(value);
    }

    /// <summary>
    ///     Change value to a value of compatible type.
    /// </summary>
    /// <param name="mapiType">Type of the mapi property.</param>
    /// <param name="value">The value.</param>
    /// <returns>Compatible value.</returns>
    internal static object ChangeType(MapiPropertyType mapiType, object value)
    {
        EwsUtilities.ValidateParam(value);

        return MapiTypeConverterMap[mapiType].ChangeType(value);
    }

    /// <summary>
    ///     Converts a MAPI Integer value.
    /// </summary>
    /// <remarks>
    ///     Usually the value is an integer but there are cases where the value has been "schematized" to an
    ///     Enumeration value (e.g. NoData) which we have no choice but to fallback and represent as a string.
    /// </remarks>
    /// <param name="s">The string value.</param>
    /// <returns>Integer value or the original string if the value could not be parsed as such.</returns>
    internal static object ParseMapiIntegerValue(string s)
    {
        if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
        {
            return intValue;
        }

        return s;
    }

    /// <summary>
    ///     Determines whether MapiPropertyType is an array type.
    /// </summary>
    /// <param name="mapiType">Type of the mapi.</param>
    /// <returns>True if this is an array type.</returns>
    internal static bool IsArrayType(MapiPropertyType mapiType)
    {
        return MapiTypeConverterMap[mapiType].IsArray;
    }
}
