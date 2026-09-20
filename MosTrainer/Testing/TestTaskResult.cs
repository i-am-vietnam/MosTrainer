namespace MosTrainer.Testing
{
    internal sealed class TestTaskResult
    {
        public TestTaskResult(string projectId, string taskId, bool pass, string message)
        {
            ProjectId = projectId;
            TaskId = taskId;
            Pass = pass;
            Message = message ?? "";
        }

        public string ProjectId { get; private set; }
        public string TaskId { get; private set; }
        public bool Pass { get; private set; }
        public string Message { get; private set; }
    }
}
