using MosTrainer.Core.Models;
using MosTrainer.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MosTrainer.Projects
{
    public class ProjectLoader
    {
        private readonly ProjectValidator _validator = new ProjectValidator();

        public List<ProjectPackage> LoadAll(string rootProjectsFolder, string languageCode)
        {
            if (string.IsNullOrWhiteSpace(rootProjectsFolder) || !Directory.Exists(rootProjectsFolder))
            {
                AppLogger.Warning("ProjectLoader.LoadAll", "Projects folder does not exist: " + rootProjectsFolder);
                return new List<ProjectPackage>();
            }

            string langCode = string.IsNullOrWhiteSpace(languageCode)
                ? "en"
                : languageCode.Trim().ToLowerInvariant();

            var packages = new List<ProjectPackage>();
            string[] projectDirectories;

            try
            {
                projectDirectories = Directory.GetDirectories(rootProjectsFolder);
            }
            catch (Exception ex)
            {
                AppLogger.Error("ProjectLoader.LoadAll", "Projects folder could not be enumerated.", ex);
                return packages;
            }

            foreach (string dir in projectDirectories)
            {
                try
                {
                    ValidationResult validation = _validator.Validate(dir, langCode);
                    if (validation.HasErrors)
                    {
                        AppLogger.Error(
                            "ProjectLoader.LoadAll",
                            "Project rejected by structural validation. Errors=" + validation.ErrorCount +
                            ", Warnings=" + validation.WarningCount + ".",
                            validation.Meta == null ? new DirectoryInfo(dir).Name : validation.Meta.ProjectId);
                        continue;
                    }

                    Dictionary<string, string> lang = SelectLanguage(validation, langCode);

                    packages.Add(new ProjectPackage
                    {
                        Meta = validation.Meta,
                        Tasks = validation.Tasks,
                        Lang = lang,
                        ProjectFolderPath = dir
                    });
                }
                catch (Exception ex)
                {
                    AppLogger.Error(
                        "ProjectLoader.LoadAll",
                        "Unexpected error while loading project.",
                        ex,
                        new DirectoryInfo(dir).Name);
                    continue;
                }
            }

            return packages
                .OrderBy(p => p.Meta == null ? "" : p.Meta.ProjectId)
                .ToList();
        }

        public string GetText(ProjectPackage pkg, string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return "";

            string value;
            if (pkg != null && pkg.Lang != null && pkg.Lang.TryGetValue(key, out value))
                return value;

            return key;
        }

        private Dictionary<string, string> SelectLanguage(ValidationResult validation, string languageCode)
        {
            if (validation == null || validation.Languages == null)
                return new Dictionary<string, string>();

            Dictionary<string, string> language;
            if (validation.Languages.TryGetValue(languageCode, out language))
                return language;

            if (validation.Languages.TryGetValue("en", out language))
                return language;

            return validation.Languages.Values.FirstOrDefault() ?? new Dictionary<string, string>();
        }
    }
}
