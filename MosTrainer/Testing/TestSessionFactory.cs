using MosTrainer.Core.Models;
using System;
using System.Collections.Generic;

namespace MosTrainer.Testing
{
    internal sealed class TestSessionFactory
    {
        private readonly Random _random;

        public TestSessionFactory()
            : this(new Random())
        {
        }

        internal TestSessionFactory(Random random)
        {
            if (random == null)
                throw new ArgumentNullException("random");

            _random = random;
        }

        public TestSession Create(IEnumerable<ProjectPackage> projects, string language)
        {
            return Create(projects, language, DateTime.UtcNow);
        }

        internal TestSession Create(
            IEnumerable<ProjectPackage> projects,
            string language,
            DateTime startedAtUtc)
        {
            if (projects == null)
                throw new ArgumentNullException("projects");

            var eligibleProjects = GetEligibleProjects(projects);
            if (eligibleProjects.Count < TestSession.RequiredProjectCount)
                throw new InvalidOperationException("Testing Mode requires at least 7 valid projects.");

            Shuffle(eligibleProjects);
            var selectedProjects = eligibleProjects.GetRange(0, TestSession.RequiredProjectCount);

            return new TestSession(
                Guid.NewGuid(),
                language,
                selectedProjects,
                startedAtUtc);
        }

        private static List<ProjectPackage> GetEligibleProjects(IEnumerable<ProjectPackage> projects)
        {
            var eligibleProjects = new List<ProjectPackage>();
            var projectIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (ProjectPackage project in projects)
            {
                if (project == null ||
                    project.Meta == null ||
                    string.IsNullOrWhiteSpace(project.Meta.ProjectId) ||
                    project.Tasks == null ||
                    project.Tasks.Count == 0)
                {
                    continue;
                }

                string projectId = project.Meta.ProjectId.Trim();
                if (projectIds.Add(projectId))
                    eligibleProjects.Add(project);
            }

            return eligibleProjects;
        }

        private void Shuffle(List<ProjectPackage> projects)
        {
            for (int i = projects.Count - 1; i > 0; i--)
            {
                int swapIndex = _random.Next(i + 1);
                ProjectPackage temporary = projects[i];
                projects[i] = projects[swapIndex];
                projects[swapIndex] = temporary;
            }
        }
    }
}
