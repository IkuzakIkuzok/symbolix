
// (c) 2023 Kazuki KOHZUKI

using System;

namespace Symbolix;

#nullable enable

public partial class ThisAddIn
{
    private void ThisAddIn_Startup(object? sender, EventArgs e)
    {

    } // private void ThisAddIn_Startup(object?, EventArgs)

    private void ThisAddIn_Shutdown(object? sender, EventArgs e)
    {

    } // private void ThisAddIn_Shutdown(object?, EventArgs)

    #region VSTO generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InternalStartup()
    {
        Startup += ThisAddIn_Startup;
        Shutdown += ThisAddIn_Shutdown;
    } // private void InternalStartup ()

    #endregion
} // public partial class ThisAddIn
