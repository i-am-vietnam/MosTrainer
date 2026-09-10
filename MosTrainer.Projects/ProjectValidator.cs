using MosTrainer.Core.Diagnostics;
using MosTrainer.Core.Models;
using MosTrainer.Core.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;

namespace MosTrainer.Projects
{
    public class ProjectValidator
    {
        public ValidationResult Validate(string projectFolderPath, string languageCode = "")
        {
            ValidationResult result = new ValidationResult();
            string fallbackProjectId = GetFolderName(projectFolderPath);

            if (string.IsNullOrWhiteSpace(projectFolderPath) || !Directory.Exists(projectFolderPath))
            {
                AddIssue(result, ValidationSeverity.Error, "PROJECT_FOLDER_INVALID",
                    fallbackProjectId, "", "Project folder does not exist: " + projectFolderPath);
                return result;
            }

            string metaPath = Path.Combine(projectFolderPath, "meta.json");
            string tasksPath = Path.Combine(projectFolderPath, "tasks.json");

            if (!File.Exists(metaPath))
            {
                AddIssue(result, ValidationSeverity.Error, "META_MISSING",
                    fallbackProjectId, "", "meta.json is missing.");
            }
            else
            {
                result.Meta = ReadJson<ProjectMeta>(
                    metaPath, result, "META_JSON_INVALID", fallbackProjectId, "meta.json could not be read or deserialized.");

                if (result.Meta == null)
                {
                    AddIssue(result, ValidationSeverity.Error, "META_EMPTY",
                        fallbackProjectId, "", "meta.json deserialized to null.");
                }
            }

            string projectId = result.Meta == null ? fallbackProjectId : result.Meta.ProjectId;

            if (result.Meta != null && string.IsNullOrWhiteSpace(result.Meta.ProjectId))
            {
                result.Meta.ProjectId = fallbackProjectId;
                projectId = fallbackProjectId;
                AddIssue(result, ValidationSeverity.Warning, "PROJECT_ID_FALLBACK",
                    projectId, "", "ProjectId is empty; folder name fallback will be used.");
            }

            if (result.Meta != null)
            {
                if (!JsonPropertyHasValue(metaPath, "starter") || string.IsNullOrWhiteSpace(result.Meta.Starter))
                {
                    AddIssue(result, ValidationSeverity.Error, "STARTER_NOT_DECLARED",
                        projectId, "", "Starter workbook is not declared in meta.json.");
                }
                else
                {
                    string starterPath = Path.Combine(projectFolderPath, result.Meta.Starter);
                    if (!File.Exists(starterPath))
                    {
                        AddIssue(result, ValidationSeverity.Error, "STARTER_MISSING",
                            projectId, "", "Starter workbook does not exist: " + result.Meta.Starter);
                    }
                }
            }

            if (!File.Exists(tasksPath))
            {
                AddIssue(result, ValidationSeverity.Error, "TASKS_MISSING",
                    projectId, "", "tasks.json is missing.");
            }
            else
            {
                result.Tasks = ReadJson<List<TaskDefinition>>(
                    tasksPath, result, "TASKS_JSON_INVALID", projectId, "tasks.json could not be read or deserialized.");

                if (result.Tasks == null)
                {
                    AddIssue(result, ValidationSeverity.Error, "TASKS_EMPTY",
                        projectId, "", "tasks.json deserialized to null.");
                }
            }

            ValidateTasks(result, projectId);
            LoadAndValidateLanguages(result, projectFolderPath, projectId, languageCode);
            ValidateLanguageKeys(result, projectId);

            return result;
        }

        private void ValidateTasks(ValidationResult result, string projectId)
        {
            if (result.Tasks == null)
                return;

            HashSet<string> taskIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < result.Tasks.Count; i++)
            {
                TaskDefinition task = result.Tasks[i];
                if (task == null)
                {
                    AddIssue(result, ValidationSeverity.Error, "TASK_NULL",
                        projectId, "", "Task at index " + i + " is null.");
                    continue;
                }

                task.ProjectId = projectId;
                string taskId = (task.TaskId ?? "").Trim();

                if (string.IsNullOrWhiteSpace(taskId))
                {
                    AddIssue(result, ValidationSeverity.Error, "TASK_ID_MISSING",
                        projectId, "", "A task has an empty taskId.");
                }
                else if (!taskIds.Add(taskId))
                {
                    AddIssue(result, ValidationSeverity.Error, "TASK_ID_DUPLICATE",
                        projectId, taskId, "Duplicate taskId in project.");
                }

                RequireTaskValue(result, projectId, taskId, task.TitleKey, "TITLE_KEY_MISSING", "titleKey");
                RequireTaskValue(result, projectId, taskId, task.InstructionKey, "INSTRUCTION_KEY_MISSING", "instructionKey");
                RequireTaskValue(result, projectId, taskId, task.AssertionType, "ASSERTION_TYPE_MISSING", "assertionType");

                if (!string.IsNullOrWhiteSpace(task.AssertionType) &&
                    !GradingService.IsAssertionTypeSupported(task.AssertionType))
                {
                    AddIssue(result, ValidationSeverity.Error, "ASSERTION_UNSUPPORTED",
                        projectId, taskId, "Unsupported assertionType: " + task.AssertionType);
                }
            }
        }

        private void LoadAndValidateLanguages(
            ValidationResult result,
            string projectFolderPath,
            string projectId,
            string requestedLanguageCode)
        {
            string langFolder = Path.Combine(projectFolderPath, "lang");
            if (!Directory.Exists(langFolder))
            {
                AddIssue(result, ValidationSeverity.Error, "LANG_FOLDER_MISSING",
                    projectId, "", "Language folder is missing.");
                return;
            }

            string[] languageFiles;
            try
            {
                languageFiles = Directory.GetFiles(langFolder, "*.json", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                AddIssue(result, ValidationSeverity.Error, "LANG_FOLDER_READ_FAILED",
                    projectId, "", "Language folder could not be read.", ex);
                return;
            }

            for (int i = 0; i < languageFiles.Length; i++)
            {
                string languageCode = Path.GetFileNameWithoutExtension(languageFiles[i]);
                Dictionary<string, string> language = ReadJson<Dictionary<string, string>>(
                    languageFiles[i], result, "LANG_JSON_INVALID", projectId,
                    "Language file could not be read or deserialized: " + Path.GetFileName(languageFiles[i]));

                if (language == null)
                {
                    AddIssue(result, ValidationSeverity.Error, "LANG_EMPTY",
                        projectId, "", "Language file deserialized to null: " + Path.GetFileName(languageFiles[i]));
                    continue;
                }

                result.Languages[languageCode] = language;
            }

            if (result.Languages.Count == 0)
            {
                AddIssue(result, ValidationSeverity.Error, "LANG_NONE",
                    projectId, "", "No valid language file is available.");
                return;
            }

            if (!result.Languages.ContainsKey("en"))
                AddIssue(result, ValidationSeverity.Warning, "LANG_EN_MISSING",
                    projectId, "", "en.json is missing; another available language must be used.");

            if (!result.Languages.ContainsKey("vi"))
                AddIssue(result, ValidationSeverity.Warning, "LANG_VI_MISSING",
                    projectId, "", "vi.json is missing; language fallback will be used.");

            string requested = (requestedLanguageCode ?? "").Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(requested) &&
                !result.Languages.ContainsKey(requested) &&
                result.Languages.ContainsKey("en"))
            {
                AddIssue(result, ValidationSeverity.Warning, "LANGUAGE_FALLBACK",
                    projectId, "", requested + ".json is missing; en.json fallback will be used.");
            }
        }

        private void ValidateLanguageKeys(ValidationResult result, string projectId)
        {
            if (result.Tasks == null)
                return;

            foreach (KeyValuePair<string, Dictionary<string, string>> languageEntry in result.Languages)
            {
                Dictionary<string, string> language = languageEntry.Value;

                foreach (TaskDefinition task in result.Tasks)
                {
                    if (task == null)
                        continue;

                    ValidateLanguageKey(result, projectId, task.TaskId, languageEntry.Key,
                        language, task.TitleKey, "TITLE_KEY_NOT_FOUND");
                    ValidateLanguageKey(result, projectId, task.TaskId, languageEntry.Key,
                        language, task.InstructionKey, "INSTRUCTION_KEY_NOT_FOUND");
                }
            }
        }

        private void ValidateLanguageKey(
            ValidationResult result,
            string projectId,
            string taskId,
            string languageCode,
            Dictionary<string, string> language,
            string key,
            string code)
        {
            if (string.IsNullOrWhiteSpace(key) || language == null)
                return;

            if (!language.ContainsKey(key))
            {
                AddIssue(result, ValidationSeverity.Error, code, projectId, taskId,
                    "Language key '" + key + "' was not found in " + languageCode + ".json.");
            }
        }

        private void RequireTaskValue(
            ValidationResult result,
            string projectId,
            string taskId,
            string value,
            string code,
            string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                AddIssue(result, ValidationSeverity.Error, code,
                    projectId, taskId, fieldName + " is empty.");
            }
        }

        private T ReadJson<T>(
            string path,
            ValidationResult result,
            string code,
            string projectId,
            string message) where T : class
        {
            try
            {
                return JsonConvert.DeserializeObject<T>(File.ReadAllText(path));
            }
            catch (Exception ex)
            {
                AddIssue(result, ValidationSeverity.Error, code, projectId, "", message, ex);
                return null;
            }
        }

        private bool JsonPropertyHasValue(string path, string propertyName)
        {
            try
            {
                JObject json = JObject.Parse(File.ReadAllText(path));
                foreach (JProperty property in json.Properties())
                {
                    if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                        return !string.IsNullOrWhiteSpace(Convert.ToString(property.Value));
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private void AddIssue(
            ValidationResult result,
            ValidationSeverity severity,
            string code,
            string projectId,
            string taskId,
            string message,
            Exception exception = null)
        {
            result.Issues.Add(new ValidationIssue
            {
                Severity = severity,
                Code = code,
                ProjectId = projectId ?? "",
                TaskId = taskId ?? "",
                Message = message ?? ""
            });

            string logMessage = code + ": " + message;
            if (severity == ValidationSeverity.Error)
            {
                if (exception == null)
                    AppLogger.Error("ProjectValidator", logMessage, projectId, taskId);
                else
                    AppLogger.Error("ProjectValidator", logMessage, exception, projectId, taskId);
            }
            else
            {
                AppLogger.Warning("ProjectValidator", logMessage, projectId, taskId);
            }
        }

        private string GetFolderName(string projectFolderPath)
        {
            if (string.IsNullOrWhiteSpace(projectFolderPath))
                return "";

            try
            {
                return new DirectoryInfo(projectFolderPath).Name;
            }
            catch
            {
                return "";
            }
        }
    }
}
