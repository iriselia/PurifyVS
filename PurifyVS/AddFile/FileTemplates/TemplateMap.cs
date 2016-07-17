﻿using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace FrenchKiwi.PurifyVS
{
	static class TemplateMap
	{
		static string _folder;
		static readonly string[] _templateFiles;
		const string _defaultExt = ".txt";

		static TemplateMap()
		{
			string assembly = Assembly.GetExecutingAssembly().Location;
			_folder = Path.Combine(Path.GetDirectoryName(assembly), "AddFile/FileTemplates");
			_templateFiles = Directory.GetFiles(_folder, "*" + _defaultExt, SearchOption.AllDirectories);
		}

		public static async Task<string> GetTemplateFilePath(Project project, string file)
		{
			string extension = Path.GetExtension(file).ToLowerInvariant();
			string name = Path.GetFileName(file);
			string safeName = Path.GetFileNameWithoutExtension(file);
			string relative = PackageUtilities.MakeRelative(project.GetRootFolder(), Path.GetDirectoryName(file));

			string templateFile = null;

			// Look for direct file name matches
			if (_templateFiles.Any(f => Path.GetFileName(f).Equals(name + _defaultExt, StringComparison.OrdinalIgnoreCase)))
			{
				templateFile = GetTemplate(name);
			}

			// Look for file extension matches
			else if (_templateFiles.Any(f => Path.GetFileName(f).Equals(extension + _defaultExt, StringComparison.OrdinalIgnoreCase)))
			{
				var tmpl = AdjustForSpecific(safeName, extension);
				templateFile = GetTemplate(tmpl);
			}

			string template = await ReplaceTokens(project, safeName, relative, templateFile);
			return NormalizeLineEndings(template);
		}

		private static string GetTemplate(string name)
		{
			return Path.Combine(_folder, name + _defaultExt);
		}

		private static async Task<string> ReplaceTokens(Project project, string name, string relative, string templateFile)
		{
			if (string.IsNullOrEmpty(templateFile))
				return templateFile;

			string rootNs = project.GetRootNamespace();
			string ns = string.IsNullOrEmpty(rootNs) ? "MyNamespace" : rootNs;

			if (!string.IsNullOrEmpty(relative))
			{
				ns += "." + ProjectHelper.CleanNameSpace(relative);
			}

			using (var reader = new StreamReader(templateFile))
			{
				string content = await reader.ReadToEndAsync();

				return content.Replace("{namespace}", ns)
							  .Replace("{itemname}", name);
			}
		}

		private static string NormalizeLineEndings(string content)
		{
			if (string.IsNullOrEmpty(content))
				return content;

			return Regex.Replace(content, @"\r\n|\n\r|\n|\r", "\r\n");
		}

		private static string AdjustForSpecific(string safeName, string extension)
		{
			if (Regex.IsMatch(safeName, "^I[A-Z].*"))
				return extension += "-interface";

			return extension;
		}
	}
}
