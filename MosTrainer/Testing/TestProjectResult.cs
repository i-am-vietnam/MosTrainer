using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace MosTrainer.Testing
{
    internal sealed class TestProjectResult
    {
        private readonly ReadOnlyCollection<TestTaskResult> _taskResults;

        public TestProjectResult(string projectId, IList<TestTaskResult> taskResults)
        {
            if (string.IsNullOrWhiteSpace(projectId))
                throw new ArgumentException("ProjectId is required.", "projectId");
            if (taskResults == null)
                throw new ArgumentNullException("taskResults");
            if (taskResults.Count == 0)
                throw new ArgumentException("A project result must contain at least one task result.", "taskResults");

            ProjectId = projectId;
            _taskResults = new List<TestTaskResult>(taskResults).AsReadOnly();
        }

        public string ProjectId { get; private set; }
        public ReadOnlyCollection<TestTaskResult> TaskResults { get { return _taskResults; } }
        public int PassedTaskCount { get { return _taskResults.Count(result => result.Pass); } }
        public int TotalTaskCount { get { return _taskResults.Count; } }
    }
}
