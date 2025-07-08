using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace GitActivityExplorer
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("90dde637-e93a-4fc4-afc3-5ad89f2f705d")]
    public class GitActivityWindow : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GitActivityWindow"/> class.
        /// </summary>
        public GitActivityWindow() : base(null)
        {
            this.Caption = "GitActivityWindow";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new GitActivityWindowControl();
        }
    }
}
