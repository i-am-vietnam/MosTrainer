using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace MosTrainer.Testing
{
    internal sealed class TestSubmissionResult
    {
        private readonly ReadOnlyCollection<TestProjectResult> _projectResults;

        public TestSubmissionResult(
            TestSubmissionReason reason,
            IList<TestProjectResult> projectResults)
        {
            if (projectResults == null)
                throw new ArgumentNullException("projectResults");

            _projectResults = new List<TestProjectResult>(projectResults).AsReadOnly();
            Reason = reason;
            PassedTaskCount = _projectResults.Sum(result => result.PassedTaskCount);
            TotalTaskCount = _projectResults.Sum(result => result.TotalTaskCount);
            ExactScore = TestScoreCalculator.CalculateExact(_projectResults);
            DisplayScore = TestScoreCalculator.CalculateDisplay(
                ExactScore,
                TotalTaskCount > 0 && PassedTaskCount == TotalTaskCount);
        }

        public TestSubmissionReason Reason { get; private set; }
        public ReadOnlyCollection<TestProjectResult> ProjectResults { get { return _projectResults; } }
        public int PassedTaskCount { get; private set; }
        public int TotalTaskCount { get; private set; }
        public decimal ExactScore { get; private set; }
        public int DisplayScore { get; private set; }
    }
}
