using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Linq;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace MatchOptionsCommands
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidMatchOptionsCommandsPkgString)]
	[ProvideAutoLoad(UIContextGuids.SolutionExists)]
    [ProvideAutoLoad(UIContextGuids.NoSolution)]
    public sealed class MatchOptionsCommandsPackage : Package
    {
        private DTE dte;

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public MatchOptionsCommandsPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            this.dte = this.GetService(typeof(SDTE)) as DTE;

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                var matchCaseCommandID = new CommandID(GuidList.guidMatchOptionsCommandsCmdSet, (int)PkgCmdIDList.cmdidToggleMatchCase);
                var matchCaseMenuItem = new OleMenuCommand(MenuItemCallbackMatchCase, null, MenuItemBeforeQueryStatusMatchCase, matchCaseCommandID);
                mcs.AddCommand( matchCaseMenuItem );

                var matchWholeWordCommandID = new CommandID(GuidList.guidMatchOptionsCommandsCmdSet, (int)PkgCmdIDList.cmdidToggleMatchWholeWord);
                var matchWholeWordMenuItem = new OleMenuCommand(MenuItemCallbackMatchWholeWord, null, MenuItemBeforeQueryStatusMatchWholeWord, matchWholeWordCommandID);
                mcs.AddCommand(matchWholeWordMenuItem);
            }
        }
        #endregion

        private void MenuItemCallbackMatchCase(object sender, EventArgs e)
        {
            this.dte.Find.MatchCase = !this.dte.Find.MatchCase;
        }

        private void MenuItemBeforeQueryStatusMatchCase(object sender, EventArgs e)
        {
            SetCheckedStatus(sender, this.dte.Find.MatchCase);
        }

        private void MenuItemCallbackMatchWholeWord(object sender, EventArgs e)
        {
            this.dte.Find.MatchWholeWord = !this.dte.Find.MatchWholeWord;
        }

        private void MenuItemBeforeQueryStatusMatchWholeWord(object sender, EventArgs e)
        {
            SetCheckedStatus(sender, this.dte.Find.MatchWholeWord);
        }

        private static void SetCheckedStatus(object sender, bool shouldBeChecked)
        {
            OleMenuCommand oleMenuCommand = sender as OleMenuCommand;
            if (oleMenuCommand != null)
            {
                oleMenuCommand.Checked = shouldBeChecked;
            }
        }
    }
}
