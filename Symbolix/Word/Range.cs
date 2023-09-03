
// (c) 2023 Kazuki KOHZUKI

using ApiRange = Microsoft.Office.Interop.Word.Range;

namespace Symbolix.Word;

/// <summary>
/// A wrapper class for <see cref="ApiRange"/>
/// which represents a contiguous area in a document.
/// </summary>
internal sealed class Range
{
    private readonly ApiRange _range;

    internal static Range Selection
        => new(Globals.ThisAddIn.Application.Selection.Range);

    /// <summary>
    /// Gets a <see cref="Document"/> object associated with the specified range.
    /// </summary>
    internal Document Document => new(this._range.Document);

    /// <summary>
    /// Gets a <see cref="Find"/> object that contains the criteria for a find operation.
    /// </summary>
    internal Find Find => new(this._range.Find);

    /// <summary>
    /// Initializes a new instance of the <see cref="Range"/> class.
    /// </summary>
    /// <param name="range">An API range object to be wrapped.</param>
    internal Range(ApiRange range)
    {
        this._range = range;
    } // ctor (ApiRange range)
} // internal sealed class Range

