using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Interop;
using System.Windows.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.VCProjectEngine;
using Microsoft.VisualStudio;
using System.Windows;

namespace FrenchKiwi.PurifyVS
{
    /// <summary>
    /// Helper class that exposes all GUIDs used across VS Package.
    /// </summary>
    internal sealed partial class PackageGuids
    {
        public const string guidPurifyVSPkgString = "27dd9dea-6dd2-403e-929d-3ff20d896c5e";
        public const string guidPurifyVSCmdSetString = "32af8a17-bbbc-4c56-877e-fc6c6575a8cf";
        public static Guid guidPurifyVSPkg = new Guid(guidPurifyVSPkgString);
        public static Guid guidPurifyVSCmdSet = new Guid(guidPurifyVSCmdSetString);
    }
    /// <summary>
    /// Helper class that encapsulates all CommandIDs uses across VS Package.
    /// </summary>
    internal sealed partial class CommandIDs
    {
        public const int cmdIdAddNewFile = 0x0100;
        public const int cmdIdPurifyScript = 0x0200;
    }

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.guidPurifyVSPkgString)]
    public sealed class PurifyVS : ExtensionPointPackage
    {
        public static DTE DTE;
        SolutionEvents SolutionEvents;
        //DTEEvents DTEEvents;

        protected override void Initialize()
        {
            base.Initialize();
            DTE = GetService(typeof(DTE)) as DTE;
            Logger.Initialize(this, Vsix.Name);
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            AddFileHandler.Initialize(this, mcs);
            ConnectToEvents();
            FrenchKiwi.PurifyVS.PurifyScriptHandler.Initialize(this);
        }

        // Do nothing
        public void ConnectToEvents()
        {
            SolutionEvents = ((Events2)DTE.Events).SolutionEvents;
            SolutionEvents.Opened += new _dispSolutionEvents_OpenedEventHandler(OnSolutionLoaded);
            SolutionEvents.BeforeClosing += new _dispSolutionEvents_BeforeClosingEventHandler(OnBeforeClosing);

            /*
            string solutionName = Path.GetFileNameWithoutExtension(_dte.Solution.FullName);
            string projectName = "project.Name";

            _dte.Windows.Item(EnvDTE.Constants.vsWindowKindSolutionExplorer).Activate();
            ((DTE2)_dte).ToolWindows.SolutionExplorer.GetItem(solutionName + @"\" + projectName).Select(vsUISelectionType.vsUISelectionTypeSelect);

            _dte.ExecuteCommand("Project.UnloadProject");
            _dte.ExecuteCommand("Project.ReloadProject");
            _*/
        }

        private void OnSolutionLoaded()
        {
            FileEventHandler.Initialize(this);
        }
        private void OnBeforeClosing()
        {
            FileEventHandler.OnDisconnection();
        }
    }

}
