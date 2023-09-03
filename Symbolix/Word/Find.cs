
// (c) 2023 Kazuki KOHZUKI

using Microsoft.Office.Interop.Word;
using ApiFind = Microsoft.Office.Interop.Word.Find;

namespace Symbolix.Word;

/// <summary>
/// A wrapper class of <see cref="ApiFind"/>,
/// which represents the criteria for a find operation.
/// </summary>
internal sealed class Find
{
    private readonly ApiFind _find;

    /// <summary>
    /// Initializes a new instance of the <see cref="Find"/> class.
    /// </summary>
    /// <param name="find">An API find object to be wrapped.</param>
    internal Find(ApiFind find)
    {
        this._find = find;
    } // ctor (ApiFind find)

    /// <summary>
    /// Removes text and paragraph formatting from a selection or from the specified find and replacement object.
    /// </summary>
    internal void ClearAllFormatting()
    {
        this._find.ClearFormatting();
        this._find.Replacement.ClearFormatting();
    } // internal void ClearAllFormatting ()

    internal void ReplaceAll(string find, string replace)
        => Replace(find, replace,
                matchCase        : true,
                matchWholeWord   : true,
                matchWildCards   : false,
                matchSoundsLike  : false,
                matchAllWordForms: false,
                forward          : true,
                wrap             : FindWrap.Stop,
                format           : false,
                replaceMode      : ReplaceMode.All
           );

    /// <summary>
    /// Runs the specified find operation.
    /// </summary>
    /// <param name="find">The text to be searched for.</param>
    /// <param name="replace">The replacement text. To delete the text specified by the <paramref name="find"/> argument, use an empty string (<c>""</c>).</param>
    /// <param name="matchCase"><c>true</c> to specify that the find text be case-sensitive.
    /// Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).</param>
    /// <param name="matchWholeWord"><c>true</c> to have the find operation locate only entire words, not text that's part of a larger word.
    /// Corresponds to the Find whole words only check box in the Find and Replace dialog box.</param>
    /// <param name="matchWildCards"><c>true</c> to have the find text be a special search operator.
    /// Corresponds to the Use wildcards check box in the Find and Replace dialog box.</param>
    /// <param name="matchSoundsLike"><c>true</c> to have the find operation locate words that sound similar to the find text.
    /// Corresponds to the Sounds like check box in the Find and Replace dialog box.</param>
    /// <param name="matchAllWordForms"><c>true</c> to have the find operation locate all forms of the find text (for example, "sit" locates "sitting" and "sat").
    /// Corresponds to the Find all word forms check box in the Find and Replace dialog box.</param>
    /// <param name="forward"><c>true</c> to search forward (toward the end of the document).</param>
    /// <param name="wrap">Controls what happens if the search begins at a point other than the beginning of the document and the end of the document is reached
    /// (or vice versa if <paramref name="forward"/> is set to <c>false</c>). This argument also controls what happens if there's a selection or range and the search text
    /// isn't found in the selection or range. Can be one of the following <see cref="FindWrap"/> value.
    /// <list type="bullet">
    ///     <item>
    ///         <term><see cref="FindWrap.Stop"/></term>
    ///         <description>The find operation ends if the beginning or end of the search range is reached.</description>
    ///     </item>
    ///     <item>
    ///         <term><see cref="FindWrap.Continue"/></term>
    ///         <description>The find operation continues if the beginning or end of the search range is reached.</description>
    ///     </item>
    ///     <item>
    ///         <term><see cref="FindWrap.Ask"/></term>
    ///         <description>After searching the selection or range, Microsoft Word displays a message asking whether to search the remainder of the document.</description>
    ///     </item>
    /// </list>
    /// </param>
    /// <param name="format"><c>true</c> to have the find operation locate formatting in addition to or instead of the find text.</param>
    /// <param name="replaceMode">Specifies how many replacements are to be made: one, all, or none. Can be any <see cref="ReplaceMode"/> value.</param>
    internal void Replace(
        string find, string replace, bool matchCase,
        bool matchWholeWord, bool matchWildCards, bool matchSoundsLike,
        bool matchAllWordForms, bool forward, FindWrap wrap,
        bool format, ReplaceMode replaceMode)
    {
        ClearAllFormatting();
        this._find.Execute(
            FindText         : find,
            MatchCase        : matchCase,
            MatchWholeWord   : matchWholeWord,
            MatchWildcards   : matchWildCards,
            MatchSoundsLike  : matchSoundsLike,
            MatchAllWordForms: matchAllWordForms,
            Forward          : forward,
            Wrap             : wrap.ToWrapped<WdFindWrap>(),
            Format           : format,
            ReplaceWith      : replace,
            Replace          : replaceMode.ToWrapped<WdReplace>()
        );
    } // internal void Replace (string, string, bool, bool, bool, bool, bool, bool, FindWrap, bool, ReplaceMode)
} // internal sealed class Find
