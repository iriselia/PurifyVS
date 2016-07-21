using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.VCProjectEngine;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace FrenchKiwi.PurifyVS
{
	static class ProjectHelper
	{
		static DTE2 _dte = PurifyVS.DTE as DTE2;

		// Add stuff
		public static ProjectItem AddFile(this Project project, string file, string itemType = null)
		{
			if (project.IsKind(ProjectTypes.ASPNET_5, ProjectTypes.SSDT))
				return _dte.Solution.FindProjectItem(file);

			ProjectItem item = project.ProjectItems.AddFromFile(file);
			item.SetItemType(itemType);
			return item;
		}
		public static void AddIncludeDirectory(this VCProject VCProject, string directory)
		{
			directory = directory.TrimEnd('\\');

			IEnumerable projectConfigurations = VCProject.Configurations as IEnumerable;
			foreach (Object objectProjectConfig in projectConfigurations)
			{
				VCConfiguration vcProjectConfig = objectProjectConfig as VCConfiguration;
				IEnumerable projectTools = vcProjectConfig.Tools as IEnumerable;
				foreach (Object objectProjectTool in projectTools)
				{
					VCCLCompilerTool compilerTool = objectProjectTool as VCCLCompilerTool;
					if (compilerTool != null)
					{
						//string additionalIncludeDirs = compilerTool.AdditionalIncludeDirectories;
						compilerTool.AdditionalIncludeDirectories += ';' + directory;
						break;
					}
				}
			}
		}
		public static bool HasIncludeDirectory(this VCProject VCProject, string Directory)
		{
			Directory = Directory.TrimEnd('\\');

			IEnumerable projectConfigurations = VCProject.Configurations as IEnumerable;
			foreach (Object objectProjectConfig in projectConfigurations)
			{
				VCConfiguration vcProjectConfig = objectProjectConfig as VCConfiguration;
				IEnumerable projectTools = vcProjectConfig.Tools as IEnumerable;
				foreach (Object objectProjectTool in projectTools)
				{
					VCCLCompilerTool compilerTool = objectProjectTool as VCCLCompilerTool;
					if (compilerTool != null)
					{
						//string additionalIncludeDirs = compilerTool.AdditionalIncludeDirectories;
						List<string> IncludeDirs = compilerTool.AdditionalIncludeDirectories.Split(';').Distinct().ToList();
						string result = IncludeDirs.Find(
							delegate (string s)
							{
								return s == Directory;
							}
						);

						if (result != null)
						{
							return true;
						}
					}
				}
			}

			return false;
		}

		public static void RemoveIncludeDirectory(this VCProject VCProject, string Directory)
		{
			Directory = Directory.TrimEnd('\\');

			IEnumerable projectConfigurations = VCProject.Configurations as IEnumerable;
			foreach (Object objectProjectConfig in projectConfigurations)
			{
				VCConfiguration vcProjectConfig = objectProjectConfig as VCConfiguration;
				IEnumerable projectTools = vcProjectConfig.Tools as IEnumerable;
				foreach (Object objectProjectTool in projectTools)
				{
					VCCLCompilerTool compilerTool = objectProjectTool as VCCLCompilerTool;
					if (compilerTool != null)
					{
						//string additionalIncludeDirs = compilerTool.AdditionalIncludeDirectories;
						List<string> IncludeDirs = compilerTool.AdditionalIncludeDirectories.Split(';').Distinct().ToList();
						Directory = Path.GetFullPath(Directory);
						IncludeDirs.RemoveAll(
							delegate (string s)
							{
								return s.Contains(Directory);
							}
							);
						compilerTool.AdditionalIncludeDirectories = string.Join(";", IncludeDirs);
						break;
					}
				}
			}
		}

		public static Project GetActiveProject()
		{
			try
			{
				var activeSolutionProjects = _dte.ActiveSolutionProjects as Array;

				if (activeSolutionProjects != null && activeSolutionProjects.Length > 0)
					return activeSolutionProjects.GetValue(0) as Project;

				var doc = _dte.ActiveDocument;

				if (doc != null && !string.IsNullOrEmpty(doc.FullName))
				{
					var item = (_dte.Solution != null) ? _dte.Solution.FindProjectItem(doc.FullName) : null;

					if (item != null)
						return item.ContainingProject;
				}
			}
			catch (Exception ex)
			{
				Logger.Log("Error getting the active project" + ex);
			}

			return null;
		}
		private static IEnumerable<Project> GetChildProjects(Project parent)
		{
			try
			{
				if (!parent.IsKind(ProjectKinds.vsProjectKindSolutionFolder) && parent.Collection == null)  // Unloaded
					return Enumerable.Empty<Project>();

				if (!string.IsNullOrEmpty(parent.FullName))
					return new[] { parent };
			}
			catch (COMException)
			{
				return Enumerable.Empty<Project>();
			}

			return parent.ProjectItems
					.Cast<ProjectItem>()
					.Where(p => p.SubProject != null)
					.SelectMany(p => GetChildProjects(p.SubProject));
		}

		// Get stuff
		public static string GetRootFolder(this Project Project)
		{
			if (Project == null || string.IsNullOrEmpty(Project.FullName))
				return null;

			string fullPath;

			try
			{
				fullPath = Project.Properties.Item("FullPath").Value as string;
			}
			catch (ArgumentException)
			{
				try
				{
					// MFC projects don't have FullPath, and there seems to be no way to query existence
					fullPath = Project.Properties.Item("ProjectDirectory").Value as string;
				}
				catch (ArgumentException)
				{
					// Installer projects have a ProjectPath.
					fullPath = Project.Properties.Item("ProjectPath").Value as string;
				}
			}

			if (string.IsNullOrEmpty(fullPath))
				return File.Exists(Project.FullName) ? Path.GetDirectoryName(Project.FullName) : null;

			if (Directory.Exists(fullPath))
				return fullPath;

			if (File.Exists(fullPath))
				return Path.GetDirectoryName(fullPath);

			return null;
		}
		public static string GetDirectoryToVCProjectItemRecursive(dynamic item)
		{
			string result = "";

			Project project = item as Project;
			ProjectItem projectItem = item as ProjectItem;

			VCProject VCProject = item as VCProject;
			VCFilter VCFilter = item as VCFilter;
			VCFile VCFile = item as VCFile;
			// If it is a project, prune the //Build part of the path and return
			if (VCProject != null)
			{
				result = (VCProject.Object as Project).GetRootFolder().TrimEnd('\\');
				//List<string> wordsToRemove = "Build".Split(' ').ToList();
				List<string> tempResult = result.Split('\\').ToList();
				tempResult.Remove("Build");
				result = string.Join("\\", tempResult);
			}
			// If it is a project item, treat it as a filter and
			// Call this function recursively to find out the exact path
			else if (VCFilter != null)
			{
				string currentName = VCFilter.Name;

				// Files in these filters are in the root folder of the project
				if (VCFilter.Name == "Source Files" ||
					VCFilter.Name == "Header Files" ||
					VCFilter.Name == "Resource Files" ||
					VCFilter.Name == "Proto Files"
					)
				{
					currentName = "";
				}
				if (VCFilter.Parent != null)
				{
					if (Directory.Exists(currentName))
					{
						// Shouldn't ever happen
						//result = currentName;
						return null;
					}
					else
					{
						result = GetDirectoryToVCProjectItemRecursive(VCFilter.Parent) + '\\' + currentName;
					}
				}
			}
			else if (VCFile != null)
			{
				string currentName = VCFile.Name;

				if (VCFile.Parent != null)
				{
					result = GetDirectoryToVCProjectItemRecursive(VCFile.Parent) + '\\' + currentName;
				}
			}
			return result;
		}
		public static string GetDirectoryToSolutionExplorerItem(UIHierarchyItem item)
		{
			string result = "";

			Window2 window = _dte.ActiveWindow as Window2;

			if (window != null && window.Type == vsWindowType.vsWindowTypeDocument)
			{
				// if a document is active, use the document's containing directory
				Document doc = _dte.ActiveDocument;
				if (doc != null && !string.IsNullOrEmpty(doc.FullName))
				{
					ProjectItem docItem = _dte.Solution.FindProjectItem(doc.FullName);

					if (docItem != null && docItem.Properties != null)
					{
						string fileName = docItem.Properties.Item("FullPath").Value.ToString();
						if (File.Exists(fileName))
							return Path.GetDirectoryName(fileName);
					}
				}
			}

			dynamic ProjectItem = item.Object;
			result = GetDirectoryToVCProjectItemRecursive(ProjectItem.Object);
			return result;
		}
		public static VCFilter GetFilterToPath(this VCProject VCProject, string file)
		{
			VCFilter result = null;
			string fileDir = Path.GetDirectoryName(file);
			string fileName = Path.GetFileName(file);
			string fileExt = Path.GetExtension(fileName);

			string rootDir = (VCProject.Object as Project).GetRootFolder().TrimEnd('\\');
			List<string> tempRootDir = rootDir.Split('\\').ToList();
			tempRootDir.Remove("Build");
			rootDir = string.Join("\\", tempRootDir);

			if (!fileDir.StartsWith(rootDir))
			{
				return null;
			}

			VCProjectItem Parent = VCProject;

			if (fileDir == rootDir)
			{
				if (fileExt == ".h" || fileExt == ".hpp")
				{
					VCFilter filter = Parent.GetFilterOrCreateNewFilterInProjectOrProjectItem("Header Files");
					if (filter != null)
					{
						result = filter;
					}
				}
				else if (fileExt == ".c" || fileExt == ".cpp" || fileExt == ".c++" || fileExt == ".ipp" || fileExt == ".inl")
				{
					VCFilter filter = Parent.GetFilterOrCreateNewFilterInProjectOrProjectItem("Source Files");
					if (filter != null)
					{
						result = filter;
					}
				}
				else if (fileExt == ".proto" || fileExt == ".capnp")
				{
					VCFilter filter = Parent.GetFilterOrCreateNewFilterInProjectOrProjectItem("Proto Files");
					if (filter != null)
					{
						result = filter;
					}
				}
				else
				{
					VCFilter filter = Parent.GetFilterOrCreateNewFilterInProjectOrProjectItem("Resource Files");
					if (filter != null)
					{
						result = filter;
					}
				}
			}
			else
			{
				string filterDir = fileDir.Substring(rootDir.Length).Trim('\\');
				List<string> folders = filterDir.Split('\\').ToList();

				int count = folders.Count;
				for (int i = 0; i < count; i++)
				{
					if (folders.First() == "")
					{
						folders.RemoveAt(0);
						continue;
					}
					VCFilter filter = Parent.GetFilterOrCreateNewFilterInProjectOrProjectItem(folders.First());
					if (filter != null)
					{
						folders.RemoveAt(0);
						Parent = filter;
					}

					if (folders.Count == 0)
					{
						result = filter;
					}
				}
			}


			return result;
		}
		private static VCFilter GetFilterOrCreateNewFilterInProjectOrProjectItem(this VCProjectItem VCProjectItem, string folderName)
		{
			VCFilter result = null;
			VCProject VCProject = VCProjectItem as VCProject;
			VCFilter VCFilter = VCProjectItem as VCFilter;
			Project parentProject = null;
			ProjectItem parentItem = null;

			if (VCProject != null)
			{
				parentProject = VCProject.Object as Project;
			}
			else if (VCFilter != null)
			{
				parentItem = VCFilter.Object as ProjectItem;
			}

			dynamic parent = null;

			if (parentItem != null)
			{
				parent = parentItem;
			}
			else if (parentProject != null)
			{
				parent = parentProject;
			}

			bool foundFilter = false;
			if (parent != null)
			{
				foreach (ProjectItem projectItem in parent.ProjectItems)
				{
					VCFilter filter = projectItem.Object as VCFilter;
					if (filter != null)
					{
						if (filter.Name == folderName)
						{
							result = filter;
							foundFilter = true;
							break;
						}
					}
				}
			}

			if (!foundFilter)
			{
				if (parentProject != null)
				{
					VCProject vcProject = parentProject.Object as VCProject;
					var newFolder = vcProject.AddFilter(folderName);
					VCFilter filter = newFolder as VCFilter;
					result = filter;
				}
				else if (parentItem != null)
				{
					VCFilter parentFilter = parentItem.Object as VCFilter;
					VCFilter filter = parentFilter.AddFilter(folderName);
					result = filter;
				}
			}

			return result;

		}
		public static List<VCFile> GetAllFilesRecursive(this VCFilter VCFilter)
		{
			List<VCFile> Result = new List<VCFile>();
			foreach (VCFile i in VCFilter.Files)
			{
				Result.Add(i);
			}
			foreach (VCFilter i in VCFilter.Filters)
			{
				Result = Result.Concat(GetAllFilesRecursive(i)).ToList();
			}

			return Result;
		}

		public static List<VCFilter> GetAllFiltersRecursive(this VCFilter VCFilter)
		{
			List<VCFilter> Result = new List<VCFilter>();
			foreach (VCFilter i in VCFilter.Filters)
			{
				Result.Add(i);
			}
			foreach (VCFilter i in VCFilter.Filters)
			{
				Result.Concat(GetAllFiltersRecursive(i));
			}

			return Result;
		}

		// Misc
		public static void SetItemType(this ProjectItem item, string itemType)
		{
			try
			{
				if (item == null || item.ContainingProject == null)
					return;

				if (string.IsNullOrEmpty(itemType)
					|| item.ContainingProject.IsKind(ProjectTypes.WEBSITE_PROJECT)
					|| item.ContainingProject.IsKind(ProjectTypes.UNIVERSAL_APP))
					return;

				item.Properties.Item("ItemType").Value = itemType;
			}
			catch (Exception ex)
			{
				Logger.Log(ex);
			}
		}
		public static bool IsKind(this Project project, params string[] kindGuids)
		{
			foreach (var guid in kindGuids)
			{
				if (project.Kind.Equals(guid, StringComparison.OrdinalIgnoreCase))
					return true;
			}

			return false;
		}
		public static string GetRootNamespace(this Project project)
		{
			if (project == null)
				return null;

			string ns = project.Name ?? string.Empty;

			try
			{
				var prop = project.Properties.Item("RootNamespace");

				if (prop != null && prop.Value != null && !string.IsNullOrEmpty(prop.Value.ToString()))
					ns = prop.Value.ToString();
			}
			catch { /* Project doesn't have a root namespace */ }

			return CleanNameSpace(ns, stripPeriods: false);
		}
		public static string CleanNameSpace(string ns, bool stripPeriods = true)
		{
			if (stripPeriods)
			{
				ns = ns.Replace(".", "");
			}

			ns = ns.Replace(" ", "")
					 .Replace("-", "")
					 .Replace("\\", ".");

			return ns;
		}





	}
}
