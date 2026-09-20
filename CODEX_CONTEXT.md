# MosTrainer Technical Context

Last architecture inspection: 2026-09-20 at commit `e20cda6397efabcd43e69d1b76e08c5a1c8a2a08` on branch `master`.

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
| `MosTrainer/LoginForm.cs` | Hard-coded MVP authentication, language selection, mutually exclusive Training/Testing selection, and Training launch. Testing currently shows a Phase 1 notice. | Medium: Testing-mode entry point and return-to-login lifecycle. |
| `MosTrainer/LoginForm.Designer.cs` | Login and language controls. | Medium: Designer edits can be fragile. |
| `MosTrainer/AppMode.cs` | Process-wide application mode enum: `Training` or `Testing`. | Low. |
| `MosTrainer/AppSession.cs` | Process-wide selected language and app mode; mode defaults to `Training`. | Low. |
| `MosTrainer/Testing/TestSession.cs` | Pure in-memory Testing session: fixed project order/index, identity, 50-minute UTC deadline, remaining/expired calculations, and one-shot submission state. | Medium: authoritative exam state. |
| `MosTrainer/Testing/TestSessionFactory.cs` | Filters/deduplicates loader output, performs an injectable Fisher-Yates shuffle once, selects seven projects, and creates a session. | Medium. |
| `MosTrainer/Form1.cs` | Training orchestration: projects, tabs, assets, working workbook, timer, navigation, restart, grade. | High: preserve Training behavior. |
| `MosTrainer/Form1.Designer.cs` | Main UI controls. | Medium: planned Testing navigation/submit controls. |
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
  -> if Testing, show the next-phase notice and remain on LoginForm
  -> hide LoginForm
  -> create and Show Form1
  -> close LoginForm when Form1 closes

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

Consequences:

- Current project switching is destructive: it discards unsaved changes and overwrites the destination with the starter.
- The working path is unique by ProjectId, not by test session.
- Current code has no public save operation used by `Form1`.
- This lifecycle is acceptable only as the existing Training baseline; Testing navigation needs a separate non-destructive path and explicit save-before-close behavior.

## 8. UI and Language Flow

Login uses two EN/VI CheckBoxes with mutual exclusion and a separate pair of Training/Testing RadioButtons. Training is selected by default. A successful login stores both `AppSession.Language` and `AppSession.Mode`. Training launches the unchanged `Form1` flow; Testing currently displays a clear Phase 1 notice and remains on Login. `AppSession.Language` is read when `Form1` is constructed. Project task titles/instructions come from the selected package language. Most application chrome/status text is currently hard-coded English; login labels switch between Vietnamese and English.

The main form contains project ComboBox/Go, project info, task tabs, task Previous/Next, Restart Project, Grade Project, status, and timer. Closing `Form1` closes Excel; its `FormClosed` handler then closes the hidden `LoginForm`, ending the application. There is no return-to-login route today.

## 9. Diagnostics and Persistence

`AppLogger` writes daily logs to:

```text
%LocalAppData%/MosTrainer/Logs/MosTrainer-yyyy-MM-dd.log
```

Logging is best-effort and never interrupts application flow.

`MosTrainer.Data` contains a SQLite schema initializer and `ResultRepository.Save`, with connection string `MosDb` in `App.config`. Neither is invoked by the current Login/Form1 flow. Testing Mode should not assume database persistence exists unless it is deliberately added later.

## 10. Testing Mode

Phase 1 status: **COMPLETED / VERIFIED**. Phase 2 status: **COMPLETED / VERIFIED**. Phase 3+: **NOT IMPLEMENTED**.

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
- Sessions are in memory only. Login intentionally does not create one yet.

Remaining proposed model:

- `TestProjectState`: `ProjectPackage`, working path, visit/save state, and final task results.
- `TestTaskResult`: project/task identity and PASS/FAIL/message for final reporting.
- `TestScoreCalculator`: pure decimal calculation.
- `TestSubmissionService`: sequential workbook opening and calls to the existing `GradingService`.

Recommended workbook strategy:

```text
Documents/MosTrainer/Testing/<SessionId>/<ProjectId>/work.xlsx
```

Copy a starter only on first entry. Before Next/Previous/Submit, save the current workbook and close it. Revisiting a project opens its existing session file without copying the starter again. Keep the randomized ProjectIds fixed in `TestSession`; never rerandomize during navigation.

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
- Actual auto-submit, save, grading, scoring, and result display are not implemented in Phase 2.

Only projects with at least one task are eligible for the random pool. Deduplicate by `ProjectId`, use a one-time shuffle, take seven, and fail startup with a localized message if fewer than seven eligible projects remain.

Score using per-project fractions, not rounded per-task weights. Multiply before dividing:

```text
score = 1000m * Sum(passedTasksInProject / (decimal)taskCountInProject) / 7m
```

Keep `decimal` through calculation and round only the displayed final total. When every task passes, the sum of project fractions is exactly 7, so `1000m * 7m / 7m` is exactly `1000m`.

After result acknowledgement: mark the session ended, close Excel, clear session/mode state, dispose the main form, and show the existing LoginForm again without starting a second application message loop. The current Login/Form1 close handler must be deliberately changed for this lifecycle when implementation is authorized.

Suggested implementation phases:

1. Add `AppMode`, `AppSession.Mode`, and mutually exclusive Login mode selection; verify Training unchanged.
2. Add pure TestSession, fixed seven-project selection, and 50-minute deadline logic.
3. Add Testing-only UI state, create the session from Login, and display fixed project/countdown progress.
4. Add explicit workbook save plus session-scoped first-open/revisit navigation.
5. Add idempotent submit orchestration that reuses `GradingService.CheckTask`.
6. Add decimal scoring, bilingual confirmation/result UI, cleanup, and return to Login.
7. Run Training regression plus Testing failure, order, persistence, and real-Excel acceptance tests.
