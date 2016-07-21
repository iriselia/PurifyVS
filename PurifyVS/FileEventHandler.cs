using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Visual C# .NET
//using Microsoft.Office.Core;
using Extensibility;
using System.Runtime.InteropServices;
using EnvDTE;
using System.Windows.Forms;
using Microsoft.VisualStudio.VCProjectEngine;
using EnvDTE80;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.Shell;

namespace FrenchKiwi.PurifyVS
{
	public static class FileEventHandler : Object
	{
		static DTE _dte = PurifyVS.DTE;
		private static PurifyVS pkg;
		private static VCProjectEngine VCProjectEngine;
		//private static EnvDTE.ProjectItemsEvents ProjectItemsEvents;


		public static void Initialize(object package)
		{
			pkg = (PurifyVS)package;

			//MessageBox.Show("Connecting add-in");
			try
			{
				foreach (object i in _dte.Solution.Projects)
				{
					Project Project = i as Project;

					if (Project != null)
					{
						if (OnConnectItem(i))
                        {
                            break;
                        }
					}
					//OnConnectItem(i as ProjectItem);
				}
				//ProjectItemsEvents = (EnvDTE.ProjectItemsEvents)_dte.Events.GetObject("ProjectItemsEvents");
				//VCProjectEngineEvents = (VCProjectEngineEvents)_dte.Events.GetObject("VCProjectEngineEvents");
				//ProjectItemsEvents.ItemAdded += new _dispProjectItemsEvents_ItemAddedEventHandler(ProjectItemAdded);
				//ProjectItemsEvents.ItemRemoved += new _dispProjectItemsEvents_ItemRemovedEventHandler(ProjectItemRemoved);
				//ProjectItemsEvents.ItemRenamed += new _dispProjectItemsEvents_ItemRenamedEventHandler(ProjectItemRenamed);

			}
			catch (System.Exception)
			{
			}

            _dte.Events.BuildEvents.OnBuildBegin += new _dispBuildEvents_OnBuildBeginEventHandler(OnBuildBegin);
            _dte.Events.BuildEvents.OnBuildDone += new _dispBuildEvents_OnBuildDoneEventHandler(OnBuildDone);


        }
        private static void OnBuildBegin(vsBuildScope Scope, vsBuildAction Action)
        {
            OnDisconnection();
        }

        private static void OnBuildDone(vsBuildScope Scope, vsBuildAction Action)
        {
            foreach (object i in _dte.Solution.Projects)
            {
                Project Project = i as Project;

                if (Project != null)
                {
                    if (OnConnectItem(i))
                    {
                        break;
                    }
                }
                //OnConnectItem(i as ProjectItem);
            }
        }
        private static bool OnConnectItem(object Item)
		{

			Project Project = Item as Project;
			VCProject VCProject2 = Item as VCProject;
			SolutionFolder SolutionFolder2 = Item as SolutionFolder;
			VCFilter VCFilter = Item as VCFilter;
			ProjectItem ProjectItem = Item as ProjectItem;
			// Projects in solution folders are actually contained in ProjectItems 
			if (ProjectItem != null)
			{
				// Peel project from the projectItem wrapper
				if (ProjectItem.Object != null)
				{
					Project = ProjectItem.Object as Project;
				}

			}

			if (Project != null)
			{
				// If it actually has stuff in it...
				if (Project.Object != null)
				{
					// If it is a solution folder
					SolutionFolder SolutionFolder = Project.Object as SolutionFolder;
					bool a = Project.Kind == ProjectKinds.vsProjectKindSolutionFolder;
					if (SolutionFolder != null)
					{
						// Do nothing
					}
					// If it is a VCProject
					VCProject VCProject = Project.Object as VCProject;
					if (VCProject != null)
					{
						// We only set it an register events once
						if (VCProjectEngine != null)
						{
							return false;
						}

						VCProjectEngine = VCProject.VCProjectEngine;
						if (VCProjectEngine != null)
						{
							VCProjectEngineEvents VCProjectEngineEvents = VCProjectEngine.Events;
							VCProjectEngineEvents.ItemAdded += new _dispVCProjectEngineEvents_ItemAddedEventHandler(OnProjectItemAdded);
							VCProjectEngineEvents.ItemMoved += new _dispVCProjectEngineEvents_ItemMovedEventHandler(OnProjectItemMoved);
							VCProjectEngineEvents.ItemRemoved += new _dispVCProjectEngineEvents_ItemRemovedEventHandler(OnProjectItemRemoved);
							VCProjectEngineEvents.ItemRenamed += new _dispVCProjectEngineEvents_ItemRenamedEventHandler(OnProjectItemRenamed);
							VCProjectEngineEvents.ItemPropertyChange += new _dispVCProjectEngineEvents_ItemPropertyChangeEventHandler(OnProjectItemPropertyChange);
                            return true;
						}
					}
				}

				// Handle children
				if (Project.ProjectItems != null)
				{
					foreach (object i in Project.ProjectItems)
					{
						OnConnectItem(i);
					}
				}
			}
            return false;
		}

		public static void OnDisconnection()
		{
			//MessageBox.Show("Disconnecting add-in");
			VCProjectEngineEvents VCProjectEngineEvents = VCProjectEngine.Events;
			if (VCProjectEngineEvents != null)
			{
				VCProjectEngineEvents.ItemAdded -= new _dispVCProjectEngineEvents_ItemAddedEventHandler(OnProjectItemAdded);
				VCProjectEngineEvents.ItemMoved -= new _dispVCProjectEngineEvents_ItemMovedEventHandler(OnProjectItemMoved);
				VCProjectEngineEvents.ItemRemoved -= new _dispVCProjectEngineEvents_ItemRemovedEventHandler(OnProjectItemRemoved);
				VCProjectEngineEvents.ItemRenamed -= new _dispVCProjectEngineEvents_ItemRenamedEventHandler(OnProjectItemRenamed);
				VCProjectEngineEvents.ItemPropertyChange -= new _dispVCProjectEngineEvents_ItemPropertyChangeEventHandler(OnProjectItemPropertyChange);
			}

			VCProjectEngine = null;

		}

		public static void OnAddInsUpdate()
		{
		}

		public static void OnStartupComplete()
		{
		}

		public static void OnBeginShutdown()
		{
		}

		private static void OnProjectItemAdded(object Item, object ItemParent)
		{
			if (!_dte.Solution.IsOpen)
			{
				return;
			}

			VCProjectItem ActualVCItem = Item as VCProjectItem;
			VCProject ParentProject = ActualVCItem.project;
			VCFilter VCFilter = Item as VCFilter;
			if (VCFilter != null)
			{
				string DstPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(VCFilter).TrimEnd('\\');


				if (!System.IO.Directory.Exists(DstPath))
				{
					System.IO.Directory.CreateDirectory(DstPath);
					ParentProject.AddIncludeDirectory(DstPath);
				}
				else
				{
					IEnumerable<string> files = Directory.EnumerateFileSystemEntries(DstPath, "*", SearchOption.AllDirectories);
					foreach (string i in files)
					{
						if (!i.Contains("CMakeLists.txt"))
						{
							if (File.Exists(i))
							{
								// It is a file
								string DstRelPath = FileSystem.MakeRelativePath(DstPath, i).TrimEnd('\\');
								bool ShouldAdd = true;
								foreach (VCFile j in VCFilter.Files)
								{
									if (i == j.FullPath)
									{
										ShouldAdd = false;
									}
								}
								if (VCFilter.CanAddFile(i) && ShouldAdd)
								{
									VCFilter.AddFile(i);
								}
							}
							else
							{
								// It is a directory
								string DstRelPath = FileSystem.MakeRelativePath(DstPath, i).TrimEnd('\\');
								if (VCFilter.CanAddFilter(DstRelPath))
								{
									VCFilter.AddFilter(DstRelPath);
									ParentProject.AddIncludeDirectory(i);
								}
							}
						}
					}
				}
			}
			//MessageBox.Show("Item added");
		}

		private static void OnProjectItemMoved(object Item, object NewParent, object OldParent)
		{
			//ProjectHelper.GetDirectoryToProjectOrProjectItem(NewParent);
			VCProjectItem ActualVCItem = Item as VCProjectItem;
			VCProjectItem NewProjectItem = NewParent as VCProjectItem;
			VCProjectItem OldProjectItem = OldParent as VCProjectItem;
			VCProject NewParentProject = NewProjectItem.project as VCProject;
			VCProject OldParentProject = null;
			if (OldProjectItem != null)
			{
				OldParentProject = OldProjectItem.project as VCProject;
			}
			VCFile VCFile = Item as VCFile;
			VCFilter VCFilter = Item as VCFilter;
			//VCFileConfiguration VCFileConfiguration = VCFile.GetFileConfigurationForProjectConfiguration(NewParentProject.Configurations);

			// Moved a file
			if (VCFile != null)
			{
				string SrcFilePath = VCFile.FullPath;
				string DstFilePath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(NewProjectItem).TrimEnd('\\') + "\\" + Path.GetFileName(SrcFilePath);
				string DstRelPath = FileSystem.MakeRelativePath(NewParentProject.ProjectDirectory.TrimEnd('\\'), DstFilePath);

				// if a document is active, use the document's containing directory
				Documents docs = _dte.Documents;
				bool shouldOpen = false;
				bool shouldMove = true;
				if (SrcFilePath == DstFilePath)
				{
					// Same path, do nothing and return;
					return;
				}
				if (docs != null)
				{
					foreach (Document i in docs)
					{
						if (i.FullName == SrcFilePath)
						{
							//i.Save();
							i.Close(vsSaveChanges.vsSaveChangesPrompt);
							shouldOpen = true;
							docs = _dte.Documents;
							foreach (Document j in docs)
							{
								if (j.FullName == SrcFilePath)
								{
									shouldMove = false;
									shouldOpen = false;
								}
							}
						}
					}
					ProjectItem docItem = _dte.Solution.FindProjectItem(SrcFilePath);
				}

				if (shouldMove)
				{
					System.IO.Directory.Move(SrcFilePath, DstFilePath);
					string relpath = VCFile.RelativePath;
					VCFile.RelativePath = DstRelPath;
				}
				else
				{
					// Move it back because user cancelled
					VCFilter OriginalFilter = OldParentProject.GetFilterToPath(SrcFilePath);
					VCFile.Move(OriginalFilter);
				}

				if (shouldOpen)
				{
					VsShellUtilities.OpenDocument(pkg, DstFilePath);
				}
			}
			// Moved a filter
			else if (VCFilter != null)
			{
				VCFilter OldFilter = OldProjectItem as VCFilter;
				VCProject OldProject = OldProjectItem as VCProject;
				string OldPath = OldParent as string;
				// From another filter
				if (OldFilter != null)
				{
					string NewFilterPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(VCFilter).TrimEnd('\\');
					string OldFilterPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(OldFilter).TrimEnd('\\') + "\\" + VCFilter.Name;
					string DstRelPath = FileSystem.MakeRelativePath(NewFilterPath, OldFilterPath);
					// Find out which folder changed
					List<string> RelPathList = DstRelPath.Split('\\').ToList();
					//RelPathList = RelPathList.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
					RelPathList.RemoveAt(0);
					OnProjectItemRenamed(VCFilter, null, string.Join("\\", RelPathList));// RelPathList.First());
				}
				// From a project
				else if (OldProject != null)
				{
					string NewFilterPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(VCFilter).TrimEnd('\\');
					string OldFilterPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(OldProject).TrimEnd('\\') + "\\" + VCFilter.Name;
					string DstRelPath = FileSystem.MakeRelativePath(NewFilterPath, OldFilterPath);
					// Find out which folder changed
					List<string> RelPathList = DstRelPath.Split('\\').ToList();
					//RelPathList = RelPathList.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
					RelPathList.RemoveAt(0);
					OnProjectItemRenamed(VCFilter, null, string.Join("\\", RelPathList));// RelPathList.First());
				}
				else if (OldPath != null)
				{
					string NewFilterPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(VCFilter).TrimEnd('\\');
					string OldFilterPath = OldPath;
					string DstRelPath = FileSystem.MakeRelativePath(NewFilterPath, OldFilterPath);
					// Find out which folder changed
					List<string> RelPathList = DstRelPath.Split('\\').ToList();
					//RelPathList = RelPathList.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
					RelPathList.RemoveAt(0);
					OnProjectItemRenamed(VCFilter, null, string.Join("\\", RelPathList));// RelPathList.First());
				}
			}
			//VCFile.Remove();
			//VCFile.RemoveFile(VCFile);
			// 
			//MyTestClass MyClass1 = new MyTestClass("Hello");
			//            Type Type = VCFile.GetType();
			//            IEnumerable<MemberInfo> Members = Type.GetMembers();
			//            IEnumerable<FieldInfo> Fields = Type.GetRuntimeFields();
			//           MemberInfo a= Members.First();
			//            typeof(Type)
			//   .GetField("FullPath", BindingFlags.Instance | BindingFlags.NonPublic)
			//   .SetValue(VCFile, 567);
			//FieldInfo MyWriteableField = Fields.Where(a => Regex.IsMatch(a.Name, $"\\A<{nameof(VCFile.FullPath)}>k__BackingField\\Z")).FirstOrDefault();
			//MyWriteableField.SetValue(VCFile, "Another new value");


			//VCFilter NewFilter = NewParentProject.GetFilterToPath(DstFilePath);
			//VCFile.Move(NewFilter);
			//NewFilter.Name = NewFilter.Name + "c";
			//NewFilter.AddFile(DstFilePath);

			//MessageBox.Show("Item moved from: " + OldProjectItem.ItemName + " to: " + NewProjectItem.ItemName + ".");
			//ProjectHelper.GetDirectoryToVCProjectItemRecursive(NewProjectItem);
		}

		private static void OnProjectItemRemoved(object Item, object ItemParent)
		{
			VCProjectItem ActualVCItem = Item as VCProjectItem;
			VCProject ParentProject = ActualVCItem.project;
			VCFilter VCFilter = Item as VCFilter;
			VCFile VCFile = Item as VCFile;
			if (ActualVCItem != null)
			{
				VCProject VCProject = ActualVCItem as VCProject;
				if (VCProject != null)
				{
					// If removing project, we shouldn't do anything.
					return;
				}
				string DstPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(ActualVCItem).TrimEnd('\\');
				string ProjectPath = ProjectHelper.GetRootFolder(ParentProject.Object);
				//List<string> wordsToRemove = "Build".Split(' ').ToList();
				List<string> tempResult = ProjectPath.Split('\\').ToList();
				tempResult.Remove("Build");
				ProjectPath = string.Join("\\", tempResult).TrimEnd('\\');
				if (DstPath == ProjectPath)
				{
					// Don't process delete request for dummy filters such as
					// Source Files and Header Files
					return;
				}
				if (System.IO.Directory.Exists(DstPath))
				{
					System.IO.Directory.Delete(DstPath, true);
					ParentProject.RemoveIncludeDirectory(DstPath);
				}
			}
			//MessageBox.Show("Item removed");
		}

		private static void OnProjectItemRenamed(object Item, object ItemParent, string OldName)
		{
			string NewName = (Item as VCProjectItem).ItemName;
			VCProjectItem ActualVCItem = Item as VCProjectItem;
			VCProject ParentProject = ActualVCItem.project;
			VCFilter VCFilter = Item as VCFilter;
			if (VCFilter != null)
			{
				string NewFilterPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(VCFilter).TrimEnd('\\');
				string OldFilterPath = NewFilterPath;
				List<string> tempResult = OldFilterPath.Split('\\').ToList();
				tempResult[tempResult.Count - 1] = OldName;
				OldFilterPath = string.Join("\\", tempResult);

				if (NewFilterPath == OldFilterPath)
				{
					return;
				}

				ParentProject.RemoveIncludeDirectory(OldFilterPath);
				OnProjectItemAdded(VCFilter, null);

				dynamic AllFiles = VCFilter.Files;
				dynamic AllFilters = VCFilter.Filters;
				string SrcFilePath = null;
				string DstFilePath = null;
				string DstFolderPath = null;
				foreach (VCProjectItem i in AllFiles)
				{
					DstFilePath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(i).TrimEnd('\\');
					DstFolderPath = System.IO.Directory.GetParent(DstFilePath).FullName;
					string RelPath = FileSystem.MakeRelativePath(NewFilterPath, DstFilePath);
					SrcFilePath = OldFilterPath + '\\' + RelPath;
					if (!System.IO.Directory.Exists(DstFolderPath))
					{
						ParentProject.AddIncludeDirectory(DstFolderPath);
						System.IO.Directory.CreateDirectory(DstFolderPath);
					}

					VCFile VCFile = i as VCFile;
					if (VCFile != null)
					{
						OnProjectItemMoved(VCFile, VCFilter, OldFilterPath);
						string DstRelPath = FileSystem.MakeRelativePath(VCFile.project.ProjectDirectory.TrimEnd('\\'), DstFilePath);
						VCFile.RelativePath = DstRelPath;
					}
					//VCFilter NewFilter = ProjectHelper.GetFilterToPath(VCFile.project, DstFolderPath);
					//System.IO.Directory.Move(SrcFilePath, DstFilePath);
				}

				// Recursively move folders as well
				foreach (VCProjectItem i in AllFilters)
				{
					string DstPath = ProjectHelper.GetDirectoryToVCProjectItemRecursive(i).TrimEnd('\\');
					//DstFolderPath = System.IO.Directory.GetParent(DstPath).FullName;
					string RelPath = FileSystem.MakeRelativePath(NewFilterPath, DstPath);
					string SrcPath = OldFilterPath + '\\' + RelPath;

					OnProjectItemMoved(i, Item, OldFilterPath + '\\' + i.ItemName);
					/*
					if (!System.IO.Directory.Exists(DstPath))
					{
						System.IO.Directory.Move(SrcPath, DstPath);
						ParentProject.RemoveIncludeDirectory(SrcPath);
						ParentProject.AddIncludeDirectory(DstPath);
					}
					else
					{
						System.IO.Directory.Delete(SrcPath);
						ParentProject.RemoveIncludeDirectory(SrcPath);
					}
					*/
				}

				// Clean up old directory if possible
				if (OldFilterPath != null && !Directory.EnumerateFileSystemEntries(OldFilterPath).Any())
				{
					System.IO.Directory.Delete(OldFilterPath, false);
				}

			}

			//MessageBox.Show("Item renamed");
		}
		private static void OnProjectItemPropertyChange(object Item, object Tool, int propertyID)
		{
			//MessageBox.Show("Item Property Change");
		}
	}
}
