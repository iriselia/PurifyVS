using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.VCProjectEngine;
using System.Collections;
using System.Threading.Tasks;
using System.Windows.Interop;
using Microsoft.VisualStudio.Text;
using System.Text.RegularExpressions;
using System.ComponentModel.Design;

namespace FrenchKiwi.PurifyVS
{
	public static class AddFileHandler
	{
		static DTE2 _dte = PurifyVSPackage.DTE as DTE2;
		private static PurifyVSPackage pkg;
		private static OleMenuCommandService mcs;

		public static void Initialize(object package, object menuCommandService)
		{
			pkg = (PurifyVSPackage)package;
			mcs = (OleMenuCommandService)menuCommandService;
			// Register command service
			if (null != mcs)
			{
				CommandID menuCommandID = new CommandID(PackageGuids.guidPurifyVSCmdSet, PackageIds.cmdidMyCommand);
				var menuItem = new OleMenuCommand(AddFileHandler.MenuItemCallback, menuCommandID);
				menuItem.BeforeQueryStatus += AddFileHandler.MenuItem_BeforeQueryStatus;
				mcs.AddCommand(menuItem);
			}
		}

		// Determines if Add new file becomes clickable 
		public static void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
		{
			var button = (OleMenuCommand)sender;
			button.Visible = button.Enabled = false;

			UIHierarchyItem item = null;
			var items = (Array)_dte.ToolWindows.SolutionExplorer.SelectedItems;
			foreach (UIHierarchyItem selItem in items)
			{
				item = selItem;
			}

			if (item == null)
				return;

			var project = item.Object as Project;

			if (project == null || !project.Kind.Equals(EnvDTE.Constants.vsProjectKindSolutionItems, StringComparison.OrdinalIgnoreCase))
				button.Visible = button.Enabled = true;
		}

		private static string PromptForFileName(string folder)
		{
			DirectoryInfo dir = new DirectoryInfo(folder);
			var dialog = new FileNameDialog(dir.Name);

			var hwnd = new IntPtr(_dte.MainWindow.HWnd);
			var window = (System.Windows.Window)HwndSource.FromHwnd(hwnd).RootVisual;
			dialog.Owner = window;

			var result = dialog.ShowDialog();
			return (result.HasValue && result.Value) ? dialog.Input : string.Empty;
		}
		static string[] GetParsedInput(string input)
		{
			// var tests = new string[] { "file1.txt", "file1.txt, file2.txt", ".ignore", ".ignore.(old,new)", "license", "folder/",
			//    "folder\\", "folder\\file.txt", "folder/.thing", "page.aspx.cs", "widget-1.(html,js)", "pages\\home.(aspx, aspx.cs)",
			//    "home.(html,js), about.(html,js,css)", "backup.2016.(old, new)", "file.(txt,txt,,)", "file_@#d+|%.3-2...3^&.txt" };
			Regex pattern = new Regex(@"[,]?([^(,]*)([\.\/\\]?)[(]?((?<=[^(])[^,]*|[^)]+)[)]?");
			List<string> results = new List<string>();
			Match match = pattern.Match(input);

			while (match.Success)
			{
				// Always 4 matches w. Group[3] being the extension, extension list, folder terminator ("/" or "\"), or empty string
				string path = match.Groups[1].Value.Trim() + match.Groups[2].Value;
				string[] extensions = match.Groups[3].Value.Split(',');

				foreach (string ext in extensions)
				{
					string value = path + ext.Trim();

					// ensure "file.(txt,,txt)" or "file.txt,,file.txt,File.TXT" returns as just ["file.txt"]
					if (value != "" && !value.EndsWith(".", StringComparison.Ordinal) && !results.Contains(value, StringComparer.OrdinalIgnoreCase))
					{
						results.Add(value);
					}
				}
				match = match.NextMatch();
			}
			return results.ToArray();
		}
		private static IVsTextView GetCurrentNativeTextView()
		{
			var textManager = (IVsTextManager)ServiceProvider.GlobalProvider.GetService(typeof(SVsTextManager));

			IVsTextView activeView = null;
			ErrorHandler.ThrowOnFailure(textManager.GetActiveView(1, null, out activeView));
			return activeView;
		}

		private static IComponentModel GetComponentModel()
		{
			return (IComponentModel)PurifyVSPackage.GetGlobalService(typeof(SComponentModel));
		}
		private static IWpfTextView GetCurentTextView()
		{
			var componentModel = GetComponentModel();
			if (componentModel == null) return null;
			var editorAdapter = componentModel.GetService<IVsEditorAdaptersFactoryService>();

			return editorAdapter.GetWpfTextView(GetCurrentNativeTextView());
		}
		// Main function for handling add new file button callback
		public static async void MenuItemCallback(object sender, EventArgs e)
		{
			UIHierarchyItem item = null;
			var items = (Array)_dte.ToolWindows.SolutionExplorer.SelectedItems;

			foreach (UIHierarchyItem selItem in items)
			{
				item = selItem;
			}

			if (item == null)
				return;

			string folder = ProjectHelper.GetDirectoryToSolutionExplorerItem(item);

			if (string.IsNullOrEmpty(folder))
				return;

			Project project = ProjectHelper.GetActiveProject();
			if (project == null)
				return;

			string input = PromptForFileName(folder).TrimStart('/', '\\').Replace("/", "\\");

			if (string.IsNullOrEmpty(input))
				return;

			string[] parsedInputs = GetParsedInput(input);

			foreach (string inputItem in parsedInputs)
			{
				input = inputItem;

				if (input.EndsWith("\\", StringComparison.Ordinal))
				{
					input = input + "__dummy__";
				}

				string file = Path.Combine(folder, input);
				string dir = Path.GetDirectoryName(file);
				VCFilter dstFilter = (project.Object as VCProject).GetFilterToPath(file);
				PackageUtilities.EnsureOutputPath(dir);

				if (!File.Exists(file))
				{
					int position = await FileSystem.WriteFile(project, file);

					try
					{
						//var projectItem = project.AddFileToProject(file);
						var projectItem = dstFilter.AddFile(file);
						(project.Object as VCProject).AddIncludeDirectory(dir);
						if (file.EndsWith("__dummy__"))
						{
							projectItem?.Delete();
							continue;
						}

						VsShellUtilities.OpenDocument(pkg, file);

						// Move cursor into position
						if (position > 0)
						{
							var view = GetCurentTextView();

							if (view != null)
								view.Caret.MoveTo(new SnapshotPoint(view.TextBuffer.CurrentSnapshot, position));
						}

						_dte.ExecuteCommand("SolutionExplorer.SyncWithActiveDocument");
						_dte.ActiveDocument.Activate();
					}
					catch (Exception ex)
					{
						Logger.Log(ex);
					}
				}
				else
				{
					System.Windows.Forms.MessageBox.Show("The file '" + file + "' already exist.");
				}
			}
		}


	}

	public static class ProjectTypes
	{
		public const string ASPNET_5 = "{8BB2217D-0F2D-49D1-97BC-3654ED321F3B}";
		public const string WEBSITE_PROJECT = "{E24C65DC-7377-472B-9ABA-BC803B73C61A}";
		public const string UNIVERSAL_APP = "{262852C6-CD72-467D-83FE-5EEB1973A190}";
		public const string NODE_JS = "{9092AA53-FB77-4645-B42D-1CCCA6BD08BD}";
		public const string SSDT = "{00d1a9c2-b5f0-4af3-8072-f6c62b433612}";
	}
}
