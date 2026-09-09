using MosTrainer.Core.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MosTrainer.Projects
{
    public class ProjectLoader
    {
        public List<ProjectPackage> LoadAll(string rootProjectsFolder, string languageCode)
        {
            if (string.IsNullOrWhiteSpace(rootProjectsFolder) || !Directory.Exists(rootProjectsFolder))
                return new List<ProjectPackage>();

            string langCode = string.IsNullOrWhiteSpace(languageCode)
                ? "en"
                : languageCode.Trim().ToLowerInvariant();

            var packages = new List<ProjectPackage>();

            foreach (string dir in Directory.GetDirectories(rootProjectsFolder))
            {
                string metaPath = Path.Combine(dir, "meta.json");
                string tasksPath = Path.Combine(dir, "tasks.json");

                if (!File.Exists(metaPath) || !File.Exists(tasksPath))
                    continue;

                try
                {
                    ProjectMeta meta = JsonConvert.DeserializeObject<ProjectMeta>(File.ReadAllText(metaPath))
                                       ?? new ProjectMeta();

                    if (string.IsNullOrWhiteSpace(meta.ProjectId))
                        meta.ProjectId = new DirectoryInfo(dir).Name;

                    List<TaskDefinition> tasks = JsonConvert.DeserializeObject<List<TaskDefinition>>(File.ReadAllText(tasksPath))
                                                 ?? new List<TaskDefinition>();

                    foreach (TaskDefinition task in tasks)
                    {
                        if (task != null)
                            task.ProjectId = meta.ProjectId;
                    }

                    Dictionary<string, string> lang = LoadLanguage(dir, langCode);

                    packages.Add(new ProjectPackage
                    {
                        Meta = meta,
                        Tasks = tasks,
                        Lang = lang,
                        ProjectFolderPath = dir
                    });
                }
                catch
                {
                    // Nếu một project lỗi JSON hoặc thiếu file, bỏ qua project đó để app không crash.
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

        private Dictionary<string, string> LoadLanguage(string projectFolder, string languageCode)
        {
            string langFolder = Path.Combine(projectFolder, "lang");
            string langPath = Path.Combine(langFolder, languageCode + ".json");
            string enPath = Path.Combine(langFolder, "en.json");

            string pathToRead = null;

            if (File.Exists(langPath))
                pathToRead = langPath;
            else if (File.Exists(enPath))
                pathToRead = enPath;

            if (string.IsNullOrWhiteSpace(pathToRead))
                return new Dictionary<string, string>();

            try
            {
                return JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(pathToRead))
                       ?? new Dictionary<string, string>();
            }
            catch
            {
                return new Dictionary<string, string>();
            }
        }
    }
}
