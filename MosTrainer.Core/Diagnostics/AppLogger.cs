using System;
using System.IO;
using System.Text;

namespace MosTrainer.Core.Diagnostics
{
    public static class AppLogger
    {
        private static readonly object SyncRoot = new object();

        public static string LogDirectoryPath
        {
            get
            {
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrWhiteSpace(localAppData))
                    localAppData = Path.GetTempPath();

                return Path.Combine(localAppData, "MosTrainer", "Logs");
            }
        }

        public static void Info(string source, string message, string projectId = "", string taskId = "")
        {
            Write("INFO", source, message, projectId, taskId, null);
        }

        public static void Warning(string source, string message, string projectId = "", string taskId = "")
        {
            Write("WARNING", source, message, projectId, taskId, null);
        }

        public static void Error(string source, string message, string projectId = "", string taskId = "")
        {
            Write("ERROR", source, message, projectId, taskId, null);
        }

        public static void Error(string source, string message, Exception exception, string projectId = "", string taskId = "")
        {
            Write("ERROR", source, message, projectId, taskId, exception);
        }

        private static void Write(
            string level,
            string source,
            string message,
            string projectId,
            string taskId,
            Exception exception)
        {
            try
            {
                lock (SyncRoot)
                {
                    DateTime timestamp = DateTime.Now;
                    Directory.CreateDirectory(LogDirectoryPath);

                    string logPath = Path.Combine(
                        LogDirectoryPath,
                        "MosTrainer-" + timestamp.ToString("yyyy-MM-dd") + ".log");

                    StringBuilder line = new StringBuilder();
                    line.Append(timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    line.Append(" [").Append(Sanitize(level)).Append("]");
                    line.Append(" Source=").Append(Sanitize(source));
                    line.Append(" Project=").Append(Sanitize(projectId));
                    line.Append(" Task=").Append(Sanitize(taskId));
                    line.Append(" Message=").Append(Sanitize(message));

                    if (exception != null)
                    {
                        line.Append(" Exception=").Append(Sanitize(exception.GetType().FullName));
                        line.Append(": ").Append(Sanitize(exception.Message));

                        if (!string.IsNullOrWhiteSpace(exception.StackTrace))
                            line.Append(" StackTrace=").Append(Sanitize(exception.StackTrace));
                    }

                    File.AppendAllText(logPath, line + Environment.NewLine, new UTF8Encoding(false));
                }
            }
            catch
            {
                // Logging must never affect the application flow.
            }
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "-";

            return value
                .Replace("\r\n", "\\n")
                .Replace("\r", "\\n")
                .Replace("\n", "\\n");
        }
    }
}
