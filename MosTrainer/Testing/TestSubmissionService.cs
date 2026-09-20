using MosTrainer.Core.Models;
using MosTrainer.Core.Services;
using MosTrainer.Excel;
using MosTrainer.Services;
using System;
using System.Collections.Generic;

namespace MosTrainer.Testing
{
    internal sealed class TestSubmissionService
    {
        private readonly TestSession _session;
        private readonly TestingWorkspaceService _workspace;
        private readonly ExcelController _excel;
        private readonly GradingService _grading;

        public TestSubmissionService(
            TestSession session,
            TestingWorkspaceService workspace,
            ExcelController excel,
            GradingService grading)
        {
            _session = session ?? throw new ArgumentNullException("session");
            _workspace = workspace ?? throw new ArgumentNullException("workspace");
            _excel = excel ?? throw new ArgumentNullException("excel");
            _grading = grading ?? throw new ArgumentNullException("grading");
        }

        public TestSubmissionResult Submit(TestSubmissionReason reason)
        {
            if (!_session.IsSubmitting || _session.IsCompleted)
                throw new InvalidOperationException("The test session is not ready for submission.");

            var projectResults = new List<TestProjectResult>();

            try
            {
                for (int projectIndex = 0; projectIndex < _session.TotalProjects; projectIndex++)
                {
                    TestProjectState state = _workspace.PrepareProject(projectIndex);
                    AssetsDeployer.DeployProjectAssets(
                        state.ProjectId,
                        state.Project.ProjectFolderPath);

                    try
                    {
                        _excel.OpenWorkbook(state.WorkingWorkbookPath);
                        projectResults.Add(GradeProject(state));
                    }
                    finally
                    {
                        _excel.Close();
                    }
                }

                return new TestSubmissionResult(reason, projectResults);
            }
            catch
            {
                _excel.Close();
                throw;
            }
        }

        private TestProjectResult GradeProject(TestProjectState state)
        {
            var taskResults = new List<TestTaskResult>();
            foreach (TaskDefinition task in state.Project.Tasks)
            {
                var grade = _grading.CheckTask(task);
                taskResults.Add(new TestTaskResult(
                    state.ProjectId,
                    task.TaskId,
                    grade.pass,
                    grade.message));
            }

            return new TestProjectResult(state.ProjectId, taskResults);
        }
    }
}
