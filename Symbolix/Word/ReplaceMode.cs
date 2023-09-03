
// (c) 2023 Kazuki KOHZUKI

using Microsoft.Office.Interop.Word;

namespace Symbolix.Word;

/// <summary>
/// Specifies the number of replacements to be made when find and replace is used.
/// </summary>
internal enum ReplaceMode
{
    /// <summary>
    /// Replace no occurrences.
    /// </summary>
    [EnumWraps(WdReplace.wdReplaceNone)]
    None = 0,

    /// <summary>
    /// Replace the first occurrence encountered.
    /// </summary>
    [EnumWraps(WdReplace.wdReplaceOne)]
    One = 1,

    /// <summary>
    /// Replace all occurrences.
    /// </summary>
    [EnumWraps(WdReplace.wdReplaceAll)]
    All = 2,
} // internal enum ReplaceMode
