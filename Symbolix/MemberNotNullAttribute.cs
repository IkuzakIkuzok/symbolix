
// (c) 2023 Kazuki KOHZUKI

namespace System.Diagnostics.CodeAnalysis;

#nullable enable

/// <summary>
/// Specifies that the method or property will ensure that the listed field and property members have values that aren't <c>null</c>.
/// </summary>
[AttributeUsage(AttributeTargets.Method | AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
internal sealed class MemberNotNullAttribute : Attribute
{
    /// <summary>
    /// Gets field or property member names.
    /// </summary>
    public string[] Members { get; }

    /// <summary>
    /// Initializes the attribute with a field or property member.
    /// </summary>
    /// <param name="member">The field or property member that is promised to be non-null.</param>
    public MemberNotNullAttribute(string member)
    {
        Members = new[] { member };
    } // ctor (string)

    /// <summary>
    /// Initializes the attribute with the list of field and property members.
    /// </summary>
    /// <param name="members">The list of field and property members that are promised to be non-null.</param>
    public MemberNotNullAttribute(params string[] members)
    {
        Members = members;
    } // ctor (params string[])
} // internal sealed class MemberNotNullAttribute : Attribute
