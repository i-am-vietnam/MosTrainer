using System;
using System.Collections.Generic;

namespace MosTrainer.Testing
{
    internal static class TestScoreCalculator
    {
        public static decimal CalculateExact(IList<TestProjectResult> projectResults)
        {
            if (projectResults == null)
                throw new ArgumentNullException("projectResults");
            if (projectResults.Count != TestSession.RequiredProjectCount)
                throw new ArgumentException("A test score requires exactly 7 project results.", "projectResults");

            decimal sumOfProjectRatios = 0m;
            foreach (TestProjectResult projectResult in projectResults)
            {
                if (projectResult == null || projectResult.TotalTaskCount <= 0)
                    throw new ArgumentException("Every project result must contain tasks.", "projectResults");

                sumOfProjectRatios +=
                    projectResult.PassedTaskCount /
                    (decimal)projectResult.TotalTaskCount;
            }

            decimal exactScore =
                1000m *
                sumOfProjectRatios /
                TestSession.RequiredProjectCount;

            return Math.Max(0m, Math.Min(1000m, exactScore));
        }

        public static int CalculateDisplay(decimal exactScore, bool allTasksPassed)
        {
            decimal clampedScore = Math.Max(0m, Math.Min(1000m, exactScore));
            if (allTasksPassed)
                return 1000;

            int roundedScore = (int)Math.Round(clampedScore, 0, MidpointRounding.AwayFromZero);
            return Math.Max(0, Math.Min(999, roundedScore));
        }
    }
}
