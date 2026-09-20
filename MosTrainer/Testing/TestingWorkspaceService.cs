using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace MosTrainer.Testing
{
    internal sealed class TestingWorkspaceService
    {
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

            TestingRootDirectory = testingRootDirectory;
            SessionDirectory = Path.Combine(
                TestingRootDirectory,
                session.SessionId.ToString("N"));

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
