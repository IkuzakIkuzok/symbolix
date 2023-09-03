
// (c) 2023 Kazuki KOHZUKI

using Microsoft.Office.Interop.Word;

namespace Symbolix.Word;

/// <summary>
/// Specifies wrap behavior if a selection or range is specified for a find operation and the search text isn't found in the selection or range.
/// </summary>
internal enum FindWrap
{
    /// <summary>
    /// The find operation ends if the beginning or end of the search range is reached.
    /// </summary>
    [EnumWraps(WdFindWrap.wdFindStop)]
    Stop = 0,

    /// <summary>
    /// The find operation continues if the beginning or end of the search range is reached.
    /// </summary>
    [EnumWraps(WdFindWrap.wdFindContinue)]
    Continue = 1,

    /// <summary>
    /// After searching the selection or range, Microsoft Word displays a message asking whether to search the remainder of the document.
    /// </summary>
    [EnumWraps(WdFindWrap.wdFindAsk)]
    Ask = 2,
} // internal enum FindWrap
