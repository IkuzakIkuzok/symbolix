
// (c) 2023 Kazuki KOHZUKI

using System;

namespace Symbolix.Word;

/// <summary>
/// Specifies the wrapped value of the enum.
/// </summary>
[AttributeUsage(AttributeTargets.All, AllowMultiple = false, Inherited = false)]
internal sealed class EnumWrapsAttribute : Attribute
{
    /// <summary>
    /// The wrapped value of the enum.
    /// </summary>
    internal object Value { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="EnumWrapsAttribute{T}"/> class.
    /// </summary>
    /// <param name="value">The wrapped value of the enum.</param>
    internal EnumWrapsAttribute(object value)
    {
        this.Value = value;
    } // ctor (Enum value)
} // internal sealed class EnumWrapsAttribute<T> : Attribute
