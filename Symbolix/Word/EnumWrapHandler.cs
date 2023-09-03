
// (c) 2023 Kazuki KOHZUKI

using System;

namespace Symbolix.Word;

internal static class EnumWrapHandler
{
    internal static dynamic ToWrapped<T>(this object wrapper)
        where T : Enum
    {
        var type = wrapper.GetType();
        var fieldinfo = type.GetField(wrapper.ToString());
        if (fieldinfo is null) return default;

        var attrs = fieldinfo.GetCustomAttributes(typeof(EnumWrapsAttribute), false) as EnumWrapsAttribute[];
        return attrs[0].Value;
    } // internal static T ToWrapped<T> (this object)
} // internal static class EnumWrapHandler
