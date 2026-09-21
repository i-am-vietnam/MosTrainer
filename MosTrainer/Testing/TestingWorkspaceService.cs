using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace MosTrainer.Testing
{
    internal sealed class TestingWorkspaceService
    {
        private readonly TestSession _session;
        private readonly ReadOnlyCollection<TestProjectState> _projectStates;

        public TestingWorkspaceService(TestSession session)
            : this(
                session,
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "MosTrainer",
                    "Testing"))
        {
        }

        internal TestingWorkspaceService(TestSession session, string testingRootDirectory)
        {
            if (session == null)
                throw new ArgumentNullException("session");
            if (string.IsNullOrWhiteSpace(testingRootDirectory))
                throw new ArgumentException("Testing root directory is required.", "testingRootDirectory");

            _session = session;
            TestingRootDirectory = testingRootDirectory;
            SessionDirectory = Path.Combine(
                TestingRootDirectory,
                _session.SessionId.ToString("N"));

            var states = new List<TestProjectState>();
            foreach (var project in session.SelectedProjects)
            {
                string projectId = project.Meta.ProjectId;
                EnsureSafeProjectId(projectId);

                string workingWorkbookPath = Path.Combine(
                    SessionDirectory,
                    projectId,
                    "work.xlsx");
                states.Add(new TestProjectState(project, workingWorkbookPath));
            }

            _projectStates = states.AsReadOnly();
        }

        public string TestingRootDirectory { get; private set; }
        public string SessionDirectory { get; private set; }
        public ReadOnlyCollection<TestProjectState> ProjectStates { get { return _projectStates; } }

        public TestProjectState GetProjectState(int projectIndex)
        {
            if (projectIndex < 0 || projectIndex >= _projectStates.Count)
                throw new ArgumentOutOfRangeException("projectIndex");

            return _projectStates[projectIndex];
        }

        public TestProjectState PrepareProject(int projectIndex)
        {
            TestProjectState state = GetProjectState(projectIndex);
            string workingPath = state.WorkingWorkbookPath;

            if (File.Exists(workingPath))
            {
                state.MarkInitialized();
                return state;
            }

            if (state.IsInitialized)
            {
                throw new FileNotFoundException(
                    "Testing workbook was previously initialized but is now missing. The starter will not be recopied automatically.",
                    workingPath);
            }

            string starterPath = Path.Combine(
                state.Project.ProjectFolderPath,
                state.Project.Meta.Starter);
            if (!File.Exists(starterPath))
                throw new FileNotFoundException("Testing starter workbook was not found.", starterPath);

            string projectDirectory = Path.GetDirectoryName(workingPath);
            Directory.CreateDirectory(projectDirectory);
            File.Copy(starterPath, workingPath, false);
            state.MarkInitialized();

            return state;
        }

        // Explicitly destructive; the caller must close the owned workbook first.
        public TestProjectState ResetProject(int projectIndex)
        {
            TestProjectState state = GetProjectState(projectIndex);
            string starterPath = Path.Combine(
                state.Project.ProjectFolderPath,
                state.Project.Meta.Starter);
            if (!File.Exists(starterPath))
                throw new FileNotFoundException("Testing starter workbook was not found.", starterPath);

            string workingPath = state.WorkingWorkbookPath;
            Directory.CreateDirectory(Path.GetDirectoryName(workingPath));

            // Stage a complete copy before replacing the existing working file.
            string stagedPath = workingPath + ".reset-" + Guid.NewGuid().ToString("N");
            try
            {
                File.Copy(starterPath, stagedPath, false);
                if (File.Exists(workingPath))
                    File.Replace(stagedPath, workingPath, null);
                else
                    File.Move(stagedPath, workingPath);
                state.MarkInitialized();
                return state;
            }
            finally
            {
                if (File.Exists(stagedPath))
                {
                    try { File.Delete(stagedPath); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
        }

        public void CleanupSessionDirectory()
        {
            if (!_session.IsCompleted)
                throw new InvalidOperationException(
                    "An incomplete Testing session workspace must not be deleted.");

            string rootPath = Path.GetFullPath(TestingRootDirectory)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string sessionPath = Path.GetFullPath(SessionDirectory)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string expectedDirectoryName = _session.SessionId.ToString("N");
            string descendantPrefix = rootPath + Path.DirectorySeparatorChar;

            if (!sessionPath.StartsWith(descendantPrefix, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    Path.GetFileName(sessionPath),
                    expectedDirectoryName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The Testing session directory is outside the expected Testing root.");
            }

            if (Directory.Exists(sessionPath))
                Directory.Delete(sessionPath, true);
        }

        private static void EnsureSafeProjectId(string projectId)
        {
            if (string.IsNullOrWhiteSpace(projectId) ||
                projectId == "." ||
                projectId == ".." ||
                projectId.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                throw new InvalidOperationException("ProjectId is not safe for a Testing workspace path.");
            }
        }
    }
}
