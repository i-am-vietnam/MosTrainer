using MosTrainer.Core.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace MosTrainer.Testing
{
    internal sealed class TestSession
    {
        public const int RequiredProjectCount = 7;
        public static readonly TimeSpan DefaultDuration = TimeSpan.FromMinutes(50);

        private readonly ReadOnlyCollection<ProjectPackage> _selectedProjects;

        internal TestSession(
            Guid sessionId,
            string language,
            IList<ProjectPackage> selectedProjects,
            DateTime startedAtUtc)
        {
            if (sessionId == Guid.Empty)
                throw new ArgumentException("Session ID must not be empty.", "sessionId");
            if (string.IsNullOrWhiteSpace(language))
                throw new ArgumentException("Language is required.", "language");
            if (selectedProjects == null)
                throw new ArgumentNullException("selectedProjects");
            if (selectedProjects.Count != RequiredProjectCount)
                throw new ArgumentException("A test session must contain exactly 7 projects.", "selectedProjects");
            if (startedAtUtc.Kind != DateTimeKind.Utc)
                throw new ArgumentException("Start time must be UTC.", "startedAtUtc");

            SessionId = sessionId;
            Language = language;
            _selectedProjects = new List<ProjectPackage>(selectedProjects).AsReadOnly();
            StartedAtUtc = startedAtUtc;
            Duration = DefaultDuration;
            DeadlineUtc = StartedAtUtc.Add(Duration);
            CurrentProjectIndex = 0;
        }

        public Guid SessionId { get; private set; }
        public string Language { get; private set; }
        public ReadOnlyCollection<ProjectPackage> SelectedProjects { get { return _selectedProjects; } }
        public int CurrentProjectIndex { get; private set; }
        public int ProjectNumber { get { return CurrentProjectIndex + 1; } }
        public int TotalProjects { get { return _selectedProjects.Count; } }
        public ProjectPackage CurrentProject { get { return _selectedProjects[CurrentProjectIndex]; } }
        public DateTime StartedAtUtc { get; private set; }
        public DateTime DeadlineUtc { get; private set; }
        public TimeSpan Duration { get; private set; }
        public TimeSpan RemainingTime { get { return GetRemainingTime(DateTime.UtcNow); } }
        public bool IsExpired { get { return IsExpiredAt(DateTime.UtcNow); } }
        public bool IsSubmitting { get; private set; }
        public bool IsCompleted { get; private set; }

        public TimeSpan GetRemainingTime(DateTime currentUtc)
        {
            EnsureUtc(currentUtc, "currentUtc");

            TimeSpan remaining = DeadlineUtc - currentUtc;
            return remaining <= TimeSpan.Zero ? TimeSpan.Zero : remaining;
        }

        public bool IsExpiredAt(DateTime currentUtc)
        {
            EnsureUtc(currentUtc, "currentUtc");
            return currentUtc >= DeadlineUtc;
        }

        public void MoveToProject(int projectIndex)
        {
            if (projectIndex < 0 || projectIndex >= _selectedProjects.Count)
                throw new ArgumentOutOfRangeException("projectIndex");

            CurrentProjectIndex = projectIndex;
        }

        public bool TryBeginSubmission()
        {
            if (IsSubmitting || IsCompleted)
                return false;

            IsSubmitting = true;
            return true;
        }

        public void MarkCompleted()
        {
            if (!IsSubmitting)
                throw new InvalidOperationException("Submission has not started.");

            IsSubmitting = false;
            IsCompleted = true;
        }

        private static void EnsureUtc(DateTime value, string parameterName)
        {
            if (value.Kind != DateTimeKind.Utc)
                throw new ArgumentException("Time must be UTC.", parameterName);
        }
    }
}
