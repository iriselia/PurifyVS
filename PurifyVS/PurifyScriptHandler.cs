//------------------------------------------------------------------------------
// <copyright file="Command1.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE80;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.VisualStudio;
using System.Windows;
using System.Diagnostics;
using System.ComponentModel;
using EnvDTE;

namespace FrenchKiwi.PurifyVS
{
	internal sealed class PurifyScriptHandler
	{
		static DTE2 _dte = PurifyVS.DTE as DTE2;
		public static readonly Guid CommandSet = new Guid(PackageGuids.guidPurifyVSCmdSetString);
		private readonly Package package;

		private static string BatchScript = null;
		private static string SolutionDirectory = null;
		private static string SolutionFile = null;
		private static IVsOutputWindowPane OutputWindow;
		private static System.Diagnostics.Process proc;
		private static long LastOpened = 0;
		static int FileChangeCounter = 0;
		private static IVsStatusbar Bar;
		private static object BarIcon;

		FileSystemWatcher watcher;
		private PurifyScriptHandler(Package package)
		{
			if (package == null)
			{
				throw new ArgumentNullException("package");
			}

			this.package = package;

			OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (commandService != null)
			{
				var menuCommandID = new CommandID(CommandSet, CommandIDs.cmdIdPurifyScript);
				var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
				commandService.AddCommand(menuItem);
			}

			List<string> PathList = _dte.Solution.FullName.Split('\\').ToList();
			PathList.Remove("Build");
			PathList.RemoveAt(PathList.Count - 1);
			string Path = string.Join("\\", PathList);
			string[] BatchScriptArray = Directory.GetFiles(Path, "*GenerateProjectFiles.bat");
			SolutionDirectory = Path;
			SolutionFile = _dte.Solution.FullName;
			BatchScript = BatchScriptArray[0];


			// Create a new FileSystemWatcher and set its properties.
			watcher = new FileSystemWatcher();
			watcher.Path = SolutionDirectory + "\\Build";
			/* Watch for changes in LastAccess and LastWrite times, and
			   the renaming of files or directories. */
			watcher.NotifyFilter = NotifyFilters.LastWrite
			   | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size;
			// Only watch text files.
			watcher.Filter = "*.sln";

			// Add event handlers.
			watcher.Changed += new FileSystemEventHandler(OnChanged);
			watcher.Created += new FileSystemEventHandler(OnCreated);
			watcher.Deleted += new FileSystemEventHandler(OnRemoved);

			// Begin watching.
			watcher.EnableRaisingEvents = true;

			// Get the output window
			var outputWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

			// Ensure that the desired pane is visible
			var paneGuid = Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
			outputWindow.CreatePane(paneGuid, "General", 1, 0);
			outputWindow.GetPane(paneGuid, out OutputWindow);
		}

		public static PurifyScriptHandler Instance
		{
			get;
			private set;
		}

		private IServiceProvider ServiceProvider
		{
			get
			{
				return this.package;
			}
		}

		public static void Initialize(Package package)
		{
			Instance = new PurifyScriptHandler(package);
		}

		private void MenuItemCallback(object sender, EventArgs e)
		{
			string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);

			_dte.Solution.Close();
			File.Delete(SolutionFile);

			proc = new System.Diagnostics.Process();
			proc.StartInfo.FileName = BatchScript;
			proc.StartInfo.UseShellExecute = false;
			proc.StartInfo.CreateNoWindow = true;
			proc.StartInfo.RedirectStandardOutput = true;
			proc.StartInfo.RedirectStandardError = true;

			proc.Start();
			Bar = ServiceProvider.GetService(typeof(SVsStatusbar)) as IVsStatusbar;
			BarIcon = (short)Microsoft.VisualStudio.Shell.Interop.Constants.SBAI_Build;
			//statusBar.Animation(1, ref icon);
			Bar.Animation(1, ref BarIcon);
			Bar.SetText("Generating Project...");

			StreamReader stringBackFromProcess = proc.StandardOutput;



			// Output the message
			OutputWindow.Clear();
			//OutputWindow.OutputString("Generating Project...");

			//OutputWindow.OutputString(stringBackFromProcess.ReadToEnd());

			//Debug.Write(stringBackFromProcess.ReadToEnd());

			// or

			//Console.Write();
			//             VsShellUtilities.ShowMessageBox(
			//                 this.ServiceProvider,
			//                 stringBackFromProcess.ReadToEnd(),
			//                 title,
			//                 OLEMSGICON.OLEMSGICON_INFO,
			//                 OLEMSGBUTTON.OLEMSGBUTTON_OK,
			//                 OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
			//             IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
			// 
			//             // Get a reference to the Command window.
			//             EnvDTE.Window win = _dte.Windows.Item(EnvDTE.Constants.vsWindowKindCommandWindow);
			//             CommandWindow CW = (CommandWindow)win.Object;
			// 
			//             // Input a command into the Command window and execute it.
			//             CW.SendInput("nav http://www.microsoft.com", true);
			// 
			//             // Insert some information text into the Command window.
			//             CW.OutputString("This URL takes you to the main Microsoft website.");
			// 
			//     // Clear the contents of the Command window.
			//             MessageBox.Show("Clearing the Command window...");
			//             CW.Clear();





			// Show a message box to prove we were here
			//             VsShellUtilities.ShowMessageBox(
			//                 this.ServiceProvider,
			//                 message,
			//                 title,
			//                 OLEMSGICON.OLEMSGICON_INFO,
			//                 OLEMSGBUTTON.OLEMSGBUTTON_OK,
			//                 OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
		}
		private static void OnChanged(object source, FileSystemEventArgs e)
		{
			//MessageBox.Show("Damn son...");
			// Specify what is done when a file is changed, created, or deleted.
			Console.WriteLine("File: " + e.FullPath + " " + e.ChangeType);
			string FileExtension = Path.GetExtension(e.FullPath);
			long CurrentTime = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond / 1000;
			long ElapsedTime = CurrentTime - LastOpened;
			if (ElapsedTime > 1 && e.FullPath == SolutionFile)
			{
				//_dte.Solution.Close();
				//OutputWindow.OutputString("At least it did something");
				_dte.Solution.Open(SolutionFile);
				LastOpened = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond / 1000;
				OutputWindow.Clear();
				Bar.Animation(0, ref BarIcon);
				Bar.SetText("Ready");

			}
		}
		private static void OnCreated(object source, FileSystemEventArgs e)
		{
			//MessageBox.Show("Damn son...");
			// Specify what is done when a file is changed, created, or deleted.
			Console.WriteLine("File: " + e.FullPath + " " + e.ChangeType);
			_dte.Solution.Open(SolutionFile);

		}
		private static void OnRemoved(object source, FileSystemEventArgs e)
		{
			//MessageBox.Show("Damn son...");
			// Specify what is done when a file is changed, created, or deleted.
			Console.WriteLine("File: " + e.FullPath + " " + e.ChangeType);
			_dte.Solution.Close();

		}
	}
}
