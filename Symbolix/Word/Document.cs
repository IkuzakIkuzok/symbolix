
// (c) 2023 Kazuki KOHZUKI

using System.Runtime.InteropServices;
using ApiDoc = Microsoft.Office.Interop.Word.Document;

namespace Symbolix.Word;

/// <summary>
/// A wrapper class for <see cref="ApiDoc"/>, which represents a document.
/// </summary>
internal sealed class Document
{
    private readonly ApiDoc _document;

    /// <summary>
    /// Gets a <see cref="Document"/> object that represents the active document.
    /// </summary>
    internal static Document ActiveDocument
        => new(Globals.ThisAddIn.Application.ActiveDocument);

    /// <summary>
    /// Gets the name of the specified object.
    /// </summary>
    internal string Name => this._document.Name;

    /// <summary>
    /// Gets a <see cref="Range"/> object that represents the main document story.
    /// </summary>
    internal Range Content => new(this._document.Content);

    /// <summary>
    /// Gets the disk or Web path to the specified object.
    /// </summary>
    internal string Path => this._document.Path;

    /// <summary>
    /// Gets or sets a value that determines if changes are tracked in the specified document.
    /// </summary>
    internal bool TrackRevisions
    {
        get => this._document.TrackRevisions;
        set => this._document.TrackRevisions = value;
    }

    /// <summary>
    /// Gets or sets a value that determines if the specified document or template hasn't changed since it was last saved.
    /// </summary>
    internal bool Saved
    {
        get => this._document.Saved;
        set => this._document.Saved = value;
    }

    /// <summary>
    /// Gets a value that determines if changes to the document cannot be saved to the original document.
    /// </summary>
    internal bool ReadOnly => this._document.ReadOnly;

    /// <summary>
    /// Initializes a new instance of the <see cref="Document"/> class.
    /// </summary>
    /// <param name="document">The API document to be wrapped.</param>
    internal Document(ApiDoc document)
    {
        this._document = document;
    } // ctor (ApiDoc document)

    /// <summary>
    /// Saves the specified document. If the document hasn't been saved before, the Save As dialog box prompts the user for a file name.
    /// </summary>
    internal void Save()
    {
        try
        {
            this._document.Save();
        }
        catch (COMException e)
        {
            System.Diagnostics.Debug.WriteLine(e.Message);
        }
    } // internal void Save ()
} // internal sealed class Document
