using EnvDTE;
using Microsoft.VisualStudio.VCProjectEngine;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrenchKiwi.PurifyVS
{
	static class FileSystem
	{
		public static async Task<int> WriteFile(Project project, string file)
		{
			string extension = Path.GetExtension(file);
			string template = await TemplateMap.GetTemplateFilePath(project, file);

			var props = new Dictionary<string, string>() { { "extension", extension.ToLowerInvariant() } };

			if (!string.IsNullOrEmpty(template))
			{
				int index = template.IndexOf('$');
				template = template.Remove(index, 1);

				await WriteToDisk(file, template);
				return index;
			}

			await WriteToDisk(file, string.Empty);

			return 0;
		}

		private static async System.Threading.Tasks.Task WriteToDisk(string file, string content)
		{
			using (var writer = new StreamWriter(file, false, new System.Text.UTF8Encoding(false)))
			{
				await writer.WriteAsync(content);
			}
		}

		public static string[] CommonPath(string left, string right)
		{
			List<string> result = new List<string>();
			string[] rightArray = right.Split();
			string[] leftArray = left.Split();

			result.AddRange(rightArray.Where(r => leftArray.Any(l => l.StartsWith(r))));

			// must check other way in case left array contains smaller words than right array
			result.AddRange(leftArray.Where(l => rightArray.Any(r => r.StartsWith(l))));

			return result.Distinct().ToArray();
		}

		public static string MakeRelativePath(string workingDirectory, string fullPath)
		{
			string result = string.Empty;
			int offset;

			// this is the easy case.  The file is inside of the working directory.
			if (fullPath.StartsWith(workingDirectory))
			{
				return fullPath.Substring(workingDirectory.Length + 1);
			}

			// the hard case has to back out of the working directory
			string[] baseDirs = workingDirectory.Split(new char[] { ':', '\\', '/' });
			string[] fileDirs = fullPath.Split(new char[] { ':', '\\', '/' });

			// if we failed to split (empty strings?) or the drive letter does not match
			if (baseDirs.Length <= 0 || fileDirs.Length <= 0 || baseDirs[0] != fileDirs[0])
			{
				// can't create a relative path between separate harddrives/partitions.
				return fullPath;
			}

			// skip all leading directories that match
			for (offset = 1; offset < baseDirs.Length; offset++)
			{
				if (baseDirs[offset] != fileDirs[offset])
					break;
			}

			// back out of the working directory
			for (int i = 0; i < (baseDirs.Length - offset); i++)
			{
				result += "..\\";
			}

			// step into the file path
			for (int i = offset; i < fileDirs.Length - 1; i++)
			{
				result += fileDirs[i] + "\\";
			}

			// append the file
			result += fileDirs[fileDirs.Length - 1];

			return result;
		}

	}
}
