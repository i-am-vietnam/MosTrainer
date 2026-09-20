# Current Status

Last verified: 2026-09-20
Branch: `master`
Commit: `e20cda6397efabcd43e69d1b76e08c5a1c8a2a08` (`Da hoan thien full 21 project`)

## Stable

- C# WinForms application targeting .NET Framework 4.7.2.
- Desktop Excel workbook control and grading through Microsoft Excel Interop, with targeted OOXML checks.
- Folder-based project loading, language selection, and structural validation.
- Existing Training flow: login, choose language/project, edit workbook, navigate tasks, restart, and grade the selected task.
- Protected grading route: `GradingService.CheckTask` -> project route -> assertion route -> `IExcelController`/`ExcelController`.

## Completed

- `Excel2019_P01` through `Excel2019_P21` are present.
- 144 task definitions are present across the 21 packages.
- English and Vietnamese language files are present for all 21 packages.
- All project resources are included in the WinForms project for output deployment.
- Architecture handoff documentation exists at the repository root.
- Testing Mode Phase 1 is complete: Login offers mutually exclusive Training/Testing choices and stores the selection in `AppSession.Mode`.
- Testing Mode Phase 2 is complete: pure in-memory `TestSession`, fixed random selection of seven unique eligible projects, and a UTC-based 50-minute exam deadline.
- Testing Mode Phase 3 is complete: Testing Login creates one session, passes it to a Testing UI shell, and displays fixed project progress plus a deadline-based countdown without opening Excel.

## Verified This Inspection

- `ProjectValidator`: 21 projects, 0 errors, 0 warnings.
- Debug solution build: success, 0 reported errors.
- Current workbook lifecycle and Training flow were inspected from source.
- Testing Mode Phase 1 build and mode-selection logic: verified; Training remains the default and `Form1` is unchanged.
- Testing Mode Phase 2 selection, deduplication, insufficient-pool handling, fixed-order/index, deadline, remaining-time, expiration, and one-shot submission-state logic passed a disposable deterministic validation harness.
- Testing Mode Phase 3 Login-to-Form integration, EN/VI UI, project progress, countdown, timeout lock, Training constructor path, and absence of Excel startup passed a disposable WinForms integration harness.
- No project JSON, starter workbook, assertion, or grading code was changed by Testing Mode Phase 1, Phase 2, or Phase 3.

## In Progress

- None.

## Planned

- Preserve Training behavior unchanged.
- Add non-destructive, session-scoped workbook persistence and project navigation.
- Disable per-task grading feedback in Testing.
- Add final submission, reuse of current grading, 1000-point calculation, result display, cleanup, and return to login.

## Testing Mode

Phase 1: **COMPLETED / VERIFIED**

Phase 2: **COMPLETED / VERIFIED**

- Pure `TestSession` and `TestSessionFactory` foundation.
- Exactly seven eligible projects, deduplicated by case-insensitive ProjectId, shuffled once, and held in fixed order.
- One 50-minute real-time deadline for the full exam, derived from `StartedAtUtc` and exposed through clamped remaining-time/expiration logic.
- `IsSubmitting`/`IsCompleted` state and `TryBeginSubmission` guard prepare timeout/manual submission for one shared, one-shot pipeline.

Phase 3: **COMPLETED / VERIFIED**

- Testing Login loads current project packages, creates one `TestSession`, and passes that same instance explicitly to `Form1`.
- Testing UI shows current project identity, Project 1/7, and an `MM:SS` countdown recalculated from `DeadlineUtc` every tick.
- Training selection/Go/Grade/Restart controls are hidden and task navigation is disabled in the Testing shell.
- At timeout the UI shows `00:00`, stops the timer, locks interactions, and displays a bilingual expiration status exactly once without marking the session as submitting.
- Testing opens no workbook and creates no working copy in this phase.

Phase 4+: **NOT IMPLEMENTED**

No Testing workbook initialization/navigation/persistence, submission, grading, or score calculation has been implemented.

## Next Recommended Task

After user review of Phase 3, implement Phase 4 only: session-scoped Testing working directories, first workbook initialization, save/close before project switches, Next/Previous Project, and reopen existing session workbooks without resetting them.
