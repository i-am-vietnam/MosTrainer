# MosTrainer Technical Context

Last architecture inspection: 2026-09-21 at commit `8a4d1df4c57d7f9dfee7bd341e31067deeae89af` on branch `master`.

## 1. Project Overview

MosTrainer is an offline MOS Excel 2019 practice application for Windows. It is a C# WinForms solution targeting .NET Framework 4.7.2. Each exercise project packages a starter workbook, task definitions, metadata, localized instructions, and optional import assets. The learner edits a real workbook in desktop Microsoft Excel; grading inspects the live workbook's actual state through Microsoft Excel Interop, with OOXML inspection for features that Interop cannot reliably expose.

Current package inventory: `Excel2019_P01` through `Excel2019_P21`, 144 tasks total. On 2026-09-20 all 21 packages passed `ProjectValidator` with 0 errors and 0 warnings, and the Debug solution build succeeded.

## 2. Solution Architecture

Solution: `MosTrainer/MosTrainer.slnx`

| Project | Responsibility |
| --- | --- |
| `MosTrainer` | WinForms executable, login/main UI, application session, project assets, and embedded project content. |
| `MosTrainer.Core` | Shared models, `IExcelController`, logging, and protected `GradingService`. |
| `MosTrainer.Excel` | Excel process/workbook lifecycle and all assertion implementations using Interop/OOXML. |
| `MosTrainer.Projects` | Project discovery, JSON loading, package validation, language selection/fallback. |
| `MosTrainer.Data` | SQLite initializer/result repository prototype; currently not called by the UI flow. |

All projects target `.NETFramework,Version=v4.7.2`. The code contains a C# 7.3 compatibility comment, but the project files do not explicitly set `LangVersion`.

Primary dependencies are Microsoft Office/Excel Interop, Newtonsoft.Json 13.0.4, System.Xml.Linq/ZIP APIs, Dapper 2.1.66, and SQLite. The application requires Windows and an installed compatible desktop Excel/Office environment.

## 3. Important Files

| File | Responsibility | Risk |
| --- | --- | --- |
| `MosTrainer/Program.cs` | Application entry point; cleanup hooks; runs `LoginForm`. | Medium: owns application lifetime. |
| `MosTrainer/LoginForm.cs` | Hard-coded MVP authentication and language/mode selection. Training opens the existing Form1 path; Testing loads packages, creates one TestSession, and passes it to Form1. | Medium: Testing-mode entry point and return-to-login lifecycle. |
| `MosTrainer/LoginForm.Designer.cs` | Login and language controls. | Medium: Designer edits can be fragile. |
| `MosTrainer/AppMode.cs` | Process-wide application mode enum: `Training` or `Testing`. | Low. |
| `MosTrainer/AppSession.cs` | Process-wide selected language and app mode; mode defaults to `Training`. | Low. |
| `MosTrainer/Testing/TestSession.cs` | Pure in-memory Testing session: fixed project order/index, identity, 50-minute UTC deadline, remaining/expired calculations, and one-shot submission state. | Medium: authoritative exam state. |
| `MosTrainer/Testing/TestSessionFactory.cs` | Filters/deduplicates loader output, performs an injectable Fisher-Yates shuffle once, selects seven projects, and creates a session. | Medium. |
| `MosTrainer/Testing/TestProjectState.cs` | One selected project's package/path plus copy-once initialization state; never stores COM objects. | Medium. |
| `MosTrainer/Testing/TestingWorkspaceService.cs` | Deterministic session/project paths and non-overwriting first-visit workbook initialization. | High: protects learner work. |
| `MosTrainer/Testing/TestSubmissionService.cs` | Shared synchronous seven-workbook submission pipeline using the existing ExcelController and GradingService instances. | High. |
| `MosTrainer/Testing/TestTaskResult.cs`, `TestProjectResult.cs`, `TestSubmissionResult.cs` | Pure internal Testing result hierarchy; no COM objects and no task-result UI leakage. | Medium. |
| `MosTrainer/Testing/TestScoreCalculator.cs` | Pure decimal equal-project-weight scoring and final display rounding. | Medium. |
| `MosTrainer/Form1.cs` | Preserves Training orchestration and hosts Testing countdown, real workbook/task UI, project navigation, and timeout save/close behavior. | High: keep mode branches isolated. |
| `MosTrainer/Form1.Designer.cs` | Main UI controls, including separate Testing Previous/Next Project buttons. | Medium. |
| `MosTrainer.Projects/ProjectLoader.cs` | Loads only structurally valid project folders, selects requested language with fallback, sorts by ProjectId. | Medium. |
| `MosTrainer.Projects/ProjectValidator.cs` | Validates package files, JSON, task IDs, supported assertions, and language keys. | High: coupled to grading assertion support. |
| `MosTrainer.Core/Models/ProjectPackage.cs` | Runtime aggregate of metadata, tasks, selected language, and folder path. | Low. |
| `MosTrainer.Core/Models/ProjectMeta.cs` | `ProjectId`, `OfficeVersion`, `Version`, and starter filename. | Low. |
| `MosTrainer.Core/Models/TaskDefinition.cs` | Shared JSON task schema and assertion parameters. | High: used by every project/grader. |
| `MosTrainer.Core/Services/GradingService.cs` | Protected project/assertion routing and exception boundary. | Critical/protected. |
| `MosTrainer.Core/Interfaces/IExcelController.cs` | Assertion contract consumed by `GradingService`. | Critical/protected. |
| `MosTrainer.Excel/ExcelController.cs` | Excel lifecycle plus all assertion implementations. | Critical/protected; large COM-sensitive file. |
| `MosTrainer.Excel/ExcelSession.cs` | Current Excel Application, Workbook, PID, path, and start time. | High: exactly one active workbook session. |
| `MosTrainer/Services/AssetsDeployer.cs` | Copies optional project assets to Documents and cleans them on application exit. | Medium. |
| `MosTrainer.Core/Diagnostics/AppLogger.cs` | Best-effort daily text logs under LocalAppData. | Low. |
| `MosTrainer.Data/ResultRepository.cs` | Writes individual task results to SQLite; not wired to current UI. | Medium/deferred. |

## 4. Current Application Flow

```text
Program.Main
  -> register AssetsDeployer.CleanupAll on process/application exit
  -> Application.Run(new LoginForm())

LoginForm
  -> choose EN or VI (two CheckBoxes with mutual exclusion)
  -> choose Training or Testing (two RadioButtons; Training is default)
  -> validate hard-coded admin / 123456 credentials
  -> AppSession.Language = "en" or "vi"
  -> AppSession.Mode = Training or Testing
  -> if Training, create parameterless Form1 and preserve the existing close-to-exit behavior
  -> if Testing, ProjectLoader.LoadAll -> TestSessionFactory.Create -> new Form1(testSession)
  -> hide LoginForm and show Form1
  -> manual Submit on Project 7/7 runs the shared grading/score pipeline
  -> or DeadlineUtc expiry runs that same pipeline immediately without confirmation
  -> result OK closes Testing Form1 and shows the existing LoginForm

Form1.MainForm_Load
  -> load <application base>/Projects through ProjectLoader
  -> ProjectValidator rejects invalid packages
  -> bind valid ProjectPackage objects to project ComboBox

Go
  -> build localized task tabs
  -> deploy optional assets
  -> copy starter.xlsx to a working path
  -> open working workbook in a new visible Excel instance

Grade
  -> selected TaskDefinition
  -> GradingService.CheckTask(task)
  -> project route
  -> CheckByAssertion(task)
  -> IExcelController assertion
  -> PASS/FAIL status label
```

`btnPrev` and `btnNext` navigate task tabs, not projects. `btnRestart` recreates a fresh working workbook. The project timer is per load/restart flow and is display-only.

## 5. Project Package Format

Runtime packages live under `MosTrainer/Projects/Excel2019_Pxx/` and are copied to the executable output by `MosTrainer/MosTrainer.WinForms.csproj`.

```text
Projects/Excel2019_Pxx/
  meta.json
  tasks.json
  starter.xlsx
  lang/
    en.json
    vi.json
  assets/              # optional
```

`meta.json` maps to `ProjectMeta`. `tasks.json` maps to a list of `TaskDefinition`. Language JSON maps title/instruction keys to text. `ProjectValidator.Validate` requires metadata, declared starter, tasks, valid JSON, unique non-empty task IDs, supported assertion types, language files, and referenced language keys. `ProjectLoader.LoadAll` excludes packages with validation errors and falls back from the requested language to English, then to the first available language.

Current optional assets are used by P02/P04 (`Instructor.csv`) and P13 (`All Data.txt`). `AssetsDeployer` copies assets to `Documents/MOS Trainer/AssetsTemp/<ProjectId>`; `ExcelController.ImportedCsvAtCell` searches that location (and a legacy capitalization/path variant).

## 6. Grading Architecture — STABLE / PROTECTED

- `TaskDefinition.ProjectId` is assigned by `ProjectValidator` from `meta.json`.
- `TaskDefinition.AssertionType` selects the semantic grader; the remaining fields parameterize it.
- `GradingService.CheckTask` rejects null tasks, closed Excel, and unsupported projects, and converts unhandled grading exceptions into FAIL while logging them.
- `CheckExcel2019Task` routes P01-P21; each project method delegates to `CheckByAssertion`.
- `CheckByAssertion` calls an `IExcelController` method and returns PASS/FAIL.
- `ExcelController` checks the live workbook via COM. Some assertions call `SaveCopyAs` and inspect a temporary OOXML package so unsaved live state can still be graded.
- `ProjectValidator` calls `GradingService.IsAssertionTypeSupported`, so package validation and grading routing are intentionally coupled.

Testing Mode must call the same `GradingService.CheckTask` for each existing `TaskDefinition`. Do not create a copied `TestingGradingService`, modify task JSON for testing, or redefine PASS/FAIL semantics.

## 7. Workbook Lifecycle

Starter path:

```text
<application base>/Projects/<ProjectId>/<meta.starter>
```

Current Training working path:

```text
Documents/MosTrainer/Working/<ProjectId>/work.xlsx
```

`Form1.OpenProjectExcel` currently:

1. calls `_excel.Close()`;
2. creates the project working directory;
3. copies the starter over `work.xlsx` with overwrite enabled;
4. calls `_excel.OpenWorkbook(work.xlsx)`.

`ExcelController.OpenWorkbook` creates a new visible Excel Application, disables alerts/events, opens the file, and records the Excel PID. Only one workbook/session is owned at a time.

`ExcelController.CloseWorkbook` calls `Workbook.Close(false)` and therefore does **not** save. It releases COM, quits the owned application, waits, and kills only its recorded Excel PID if the process remains. `Form1.OnFormClosed`, project switching, Go, and Restart all ultimately use this close behavior. Restart intentionally starts from the original starter.

Training consequences:

- Training project opening remains destructive by design: it discards unsaved changes and overwrites the destination with the starter.
- The Training working path remains unique by ProjectId, not by test session.
- This lifecycle remains the unchanged Training baseline.

Testing uses a separate lifecycle:

```text
Documents/MosTrainer/Testing/<SessionId N>/<ProjectId>/work.xlsx
```

`TestingWorkspaceService` computes each path once. On first visit it copies the package starter with overwrite disabled. If the working file exists, revisits open it directly; if an initialized file later disappears, the service fails rather than silently resetting it from the starter. `TestProjectState` stores only `ProjectPackage`, ProjectId/task count, working path, and `IsInitialized`; COM ownership remains in `ExcelController`.

`ExcelController.SaveWorkbook` saves only the currently owned workbook and propagates failures. Testing project navigation prepares the target path, explicitly saves the current workbook, calls the existing `Close()` (`Workbook.Close(false)` remains unchanged), opens the target, and only then commits `TestSession.CurrentProjectIndex` plus task/project UI. A target-open failure keeps the old index and attempts to reopen the saved current workbook. Project switching never changes SessionId, StartedAtUtc, or DeadlineUtc.

## 8. UI and Language Flow

Login uses two EN/VI CheckBoxes with mutual exclusion and a separate pair of Training/Testing RadioButtons. Training is selected by default. A successful login stores both `AppSession.Language` and `AppSession.Mode`. Training launches the existing parameterless `Form1` flow. Testing loads packages in the selected language, creates one session, and passes the same instance through the Testing constructor. A clear localized error keeps Login active when a session cannot be created. `AppSession.Language` is read when Training Form1 is constructed. Project task titles/instructions come from the selected package language. Most Training chrome/status text remains hard-coded English; the new Testing shell/status text supports EN/VI without changing project language JSON.

The main form contains project ComboBox/Go, project info, task tabs, task Previous/Next, Restart Project, Grade Project, status, and timer. Training retains the existing lifecycle: closing Form1 closes Excel and then closes the hidden LoginForm, ending the application. Closing the Testing form safely saves the current workbook and returns to that same LoginForm. Successful manual submission also returns there after result acknowledgement; automatic completed-workspace cleanup remains deferred.

In Testing, `Form1` constructs one `TestingWorkspaceService` from the supplied session, lazily prepares the current `TestProjectState`, deploys existing project assets, opens the session-scoped workbook, and builds the existing localized task tabs. The project ComboBox, Go, Grade, and Restart stay hidden. Task Previous/Next keep their original task meaning; separate Previous/Next Project controls navigate the fixed seven-project order.

On Project 7/7, manual Submit shows a bilingual Yes/No confirmation. Yes rechecks the deadline, acquires `TestSession.TryBeginSubmission`, stops the UI timer, locks interactions, saves/closes the active workbook, and calls the shared `TestSubmissionService`. The service initializes unvisited workbooks through `TestingWorkspaceService`, opens each of the seven fixed projects sequentially using the same ExcelController instance used by GradingService, calls `GradingService.CheckTask` for every task, records results, and closes after each project without changing `CurrentProjectIndex`.

At deadline, `UpdateTestingCountdown` displays `00:00` and calls `HandleTestingTimeExpired`. The handler sets its private one-shot flag before stopping/locking and calls `BeginTestingSubmission(TestSubmissionReason.TimeExpired)` directly. It performs no separate Save/Close. Manual and timeout therefore share `TryBeginSubmission`, current workbook Save/Close, `TestSubmissionService`, score/result construction, completion, result dialog, and return-to-Login lifecycle; only manual submission shows confirmation.

## 9. Diagnostics and Persistence

`AppLogger` writes daily logs to:

```text
%LocalAppData%/MosTrainer/Logs/MosTrainer-yyyy-MM-dd.log
```

Logging is best-effort and never interrupts application flow.

`MosTrainer.Data` contains a SQLite schema initializer and `ResultRepository.Save`, with connection string `MosDb` in `App.config`. Neither is invoked by the current Login/Form1 flow. Testing Mode should not assume database persistence exists unless it is deliberately added later.

## 10. Testing Mode

Phase 1 status: **COMPLETED / VERIFIED**. Phase 2 status: **COMPLETED / VERIFIED**. Phase 3 status: **COMPLETED / VERIFIED**. Phase 4 status: **COMPLETED / VERIFIED**. Phase 5 status: **COMPLETED / VERIFIED**. Phase 6 status: **COMPLETED / VERIFIED**. Phase 7+: **NOT IMPLEMENTED**.

Implemented in Phase 1:

- `AppMode`: `Training` or `Testing`.
- `AppSession.Mode`: selected at login alongside language, defaulting to `Training`.
- Login mode selection through mutually exclusive RadioButtons.
- Temporary Testing behavior: store `Testing`, show a next-phase notice, and do not enter `Form1`.

Implemented in Phase 2:

- `TestSessionFactory` accepts the project list produced by `ProjectLoader` plus the selected language.
- Eligibility requires a non-null package/meta, non-empty ProjectId, and at least one task. ProjectIds are deduplicated case-insensitively.
- Fisher-Yates shuffles an eligible copy once; the first seven become the session's read-only fixed order. Fewer than seven fails with `Testing Mode requires at least 7 valid projects.`
- `TestSession` owns a new `Guid` SessionId, language, exactly seven selected projects, current project index, and project progress data.
- The full exam duration is one fixed 50-minute interval. `StartedAtUtc` is captured once at session creation and `DeadlineUtc = StartedAtUtc + 50 minutes`; project/task navigation never changes either value.
- Remaining time is always recalculated as `DeadlineUtc - currentUtc` and clamped to zero. Expiration begins at `currentUtc >= DeadlineUtc`; no tick counter, pause, or per-project reset exists.
- `TryBeginSubmission`, `IsSubmitting`, and `IsCompleted` provide a one-shot guard for the future shared manual/timeout submission pipeline.
- Sessions are in memory only; Login now creates a fresh session for each successful Testing login.

Implemented in Phase 3:

- Testing Login calls `ProjectLoader.LoadAll`, then `TestSessionFactory.Create`, and constructs `Form1(testSession)`. Form1 never creates or randomizes a second session.
- The parameterless Form1 constructor remains the Training path. `MainForm_Load` branches only when an explicit TestSession is present, so Testing returns before `LoadProjectsToCombo`.
- The Testing shell hides project selection, Go, Grade, and Restart; disables task tabs and task Previous/Next; and shows the current session ProjectId plus localized Project 1/7 progress.
- The existing WinForms timer is reused at 1000 ms. Each Testing tick obtains one `DateTime.UtcNow`, recalculates remaining time through `TestSession.GetRemainingTime`, and checks `IsExpiredAt`; it never decrements a counter.
- Countdown display is `MM:SS`, rounded up to the next whole second before expiry so `00:00` is reserved for the reached deadline. It is updated immediately during form load rather than waiting for the first timer tick.
- Deadline handling displays `00:00`, stops the timer, locks Testing interactions, and sets a private one-shot UI flag. It does not call `TryBeginSubmission`, grading, scoring, or any fake submission.
- No Testing starter is copied, no working path is assigned, and no Excel workbook is opened. Closing the Phase 3 Testing shell returns to Login; session resume remains unsupported.

Implemented in Phase 4:

- Testing has a session-scoped workspace and lazily initializes each selected project from its starter exactly once. Existing `work.xlsx` files are never overwritten on revisit.
- The first selected project opens automatically in real Excel, existing localized task tabs are built, and task Previous/Next are enabled without exposing Grade, Restart, project selection, or Go.
- Separate Previous/Next Project controls follow the fixed session order and use Save -> Close -> Open. Buttons are disabled at project boundaries and during synchronous switching.
- After every switch, countdown is recomputed from the unchanged UTC deadline. If the deadline passed while Excel blocked the UI, timeout handling runs immediately rather than waiting for another timer tick.
- Timeout is one-shot: display `00:00`, stop the timer, lock task/project interactions, save the active workbook, and close Excel after a successful save. A save failure leaves Excel open to avoid `Close(false)` data loss and reports an infrastructure error while MosTrainer stays locked.
- Manual Testing form closure performs a mandatory save; a save failure cancels the close. Successful closure then uses the existing Excel cleanup and returns to Login.
- Workbooks for unvisited projects are intentionally not created yet. The submission phase must initialize those as untouched copies before grading.
- Actual manual/automatic submission, `TryBeginSubmission`, grading, scoring, results, and session cleanup remain unimplemented.

Implemented in Phase 5:

- A separate Submit Test/Nộp bài button is hidden in Training, visible in Testing, disabled on Projects 1-6, and enabled on Project 7 when the session is active.
- Confirmation No leaves the timer/session/workbook unchanged. Confirmation Yes rechecks `DeadlineUtc`, then calls the shared `BeginTestingSubmission(TestSubmissionReason.Manual)` entry point.
- `TestSession.AbortSubmission` clears only `IsSubmitting` for infrastructure-failure retry and refuses to roll back a completed session. `TryBeginSubmission` remains the authoritative double-submit guard.
- `TestSubmissionService` prepares all seven fixed project states, including untouched copies for unvisited projects, deploys existing assets, opens each working workbook, calls the protected existing `GradingService.CheckTask` once for every task, captures results, and closes Excel in `finally` paths.
- `TestTaskResult` stores ProjectId, TaskId, Pass, and diagnostic Message. `TestProjectResult` aggregates task counts. `TestSubmissionResult` aggregates seven projects, total counts, exact decimal score, display score, and submission reason.
- `TestScoreCalculator` computes `1000m * Sum(passed / (decimal)total per project) / 7m`. It never rounds task/project fractions; display rounding occurs once. All-pass is exactly 1000, and non-perfect display is capped at 999.
- Completion is marked only after all seven workbooks are graded and the score is constructed. The bilingual result dialog displays only total score; OK closes Testing Form1 and the existing LoginForm lifecycle shows Login again.
- Infrastructure failure shows no score, calls `AbortSubmission`, and, while before deadline, reopens the current working workbook and resumes the timer from the unchanged deadline. A post-deadline failure stays locked.
- Timeout retains Phase 4 behavior and does not yet invoke submission. Completed session directories are retained.

Implemented in Phase 6:

- Deadline expiry from any project sets `00:00`, sets `_testingTimeoutHandled` before submission, stops the timer, locks Testing interactions, and calls `BeginTestingSubmission(TimeExpired)` without a confirmation dialog.
- `HandleTestingTimeExpired` no longer saves or closes Excel. Save/Close occurs exactly once in the shared submission entry point before the unchanged seven-project `TestSubmissionService` pipeline.
- Timeout results carry `TestSubmissionReason.TimeExpired`; manual results remain `Manual`. Both use the same grading, decimal score, bilingual result dialog, and existing LoginForm return lifecycle.
- Unvisited projects are initialized through the existing copy-once workspace service and graded as untouched workbooks. Backend grading never changes the user's current project index.
- Timeout infrastructure failure calls `AbortSubmission`, produces no result/score, leaves the timer stopped at `00:00`, and keeps task/project/Submit controls locked. Repeated timeout calls return through the UI one-shot guard.
- A completed timeout session returns to Login; a later Testing login creates a new SessionId and fresh 50-minute deadline.

The implemented workbook strategy is copy-once and lazy: before Previous/Next/Submit, save and close the current workbook; revisiting opens the same session file. Submission reuses these paths and initializes any unvisited project as an untouched working copy.

Recommended final grading strategy is submit-time sequential grading (Option B): save/close the current workbook, then for each of the seven fixed project states open its saved workbook, call `GradingService.CheckTask` for every task, record results, and close before the next workbook. This avoids stale hidden results when a learner revisits a project. It is slower than grading on every transition but is simpler, authoritative at submission, and keeps grading logic unchanged.

Option A (grade silently on every project transition) gives earlier results and shorter submit time, but results become stale if a learner revisits a workbook. It also invokes COM grading more often and requires careful invalidation/regrading. Option B is therefore the recommended first implementation.

Testing UI rules:

- Hide or disable project selection, Go, per-task Grade, and PASS/FAIL feedback.
- Preserve task tabs and task navigation.
- Add separate project navigation/progress and Submit Test controls.
- Require confirmation and make submission idempotent.
- Do not change Training control behavior.

Testing timer and timeout invariants:

- The future UI displays the full-session countdown as `MM:SS` from `50:00` through `00:00`. A WinForms timer may tick every 1000 ms, but every display update must query the session deadline; tick count is never authoritative.
- UI lag, focus changes, workbook loading, project/task navigation, and other open dialogs do not pause or extend the deadline.
- At `00:00`, regardless of the current project or whether all seven were visited, the UI must lock Testing navigation and start automatic submission as soon as possible.
- Timeout submission has no confirmation. Manual submission may confirm, but both must invoke one shared submission pipeline, distinguished by a reason only when that phase needs it.
- The `TryBeginSubmission` guard must be acquired before invoking that pipeline so repeated timer ticks cannot submit more than once.
- Phase 6 routes timeout through the same `BeginTestingSubmission(TimeExpired)` entry point without confirmation. There is one grading/scoring pipeline for both reasons.

Only projects with at least one task are eligible for the random pool. Deduplicate by `ProjectId`, use a one-time shuffle, take seven, and fail startup with a localized message if fewer than seven eligible projects remain.

Score using per-project fractions, not rounded per-task weights. Multiply before dividing:

```text
score = 1000m * Sum(passedTasksInProject / (decimal)taskCountInProject) / 7m
```

Keep `decimal` through calculation and round only the displayed final total. When every task passes, the sum of project fractions is exactly 7, so `1000m * 7m / 7m` is exactly `1000m`.

After result acknowledgement, Phase 5 closes Testing Form1 and the existing LoginForm close handler shows the same Login instance without starting a second application message loop. Session workspace cleanup and explicit mode reset remain deferred.

Suggested implementation phases:

1. Add `AppMode`, `AppSession.Mode`, and mutually exclusive Login mode selection; verify Training unchanged.
2. Add pure TestSession, fixed seven-project selection, and 50-minute deadline logic.
3. Connect Testing Login to one TestSession and add the Testing UI shell with fixed project/deadline progress.
4. Add explicit workbook save plus session-scoped first-open/revisit navigation.
5. Add idempotent manual submit orchestration, decimal scoring, bilingual result UI, failure rollback, and return to Login while reusing `GradingService.CheckTask`.
6. Route timeout into the same submission pipeline without confirmation.
7. Add the chosen completed-workspace retention/cleanup policy, final UI polish, and broad Training/Testing acceptance tests.
