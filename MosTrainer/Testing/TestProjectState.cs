using MosTrainer.Core.Models;

namespace MosTrainer.Testing
{
    internal sealed class TestProjectState
    {
        internal TestProjectState(ProjectPackage project, string workingWorkbookPath)
        {
            Project = project;
            WorkingWorkbookPath = workingWorkbookPath;
        }

        public ProjectPackage Project { get; private set; }
        public string ProjectId { get { return Project.Meta.ProjectId; } }
        public int TaskCount { get { return Project.Tasks.Count; } }
        public string WorkingWorkbookPath { get; private set; }
        public bool IsInitialized { get; private set; }

        internal void MarkInitialized()
        {
            IsInitialized = true;
        }
    }
}
