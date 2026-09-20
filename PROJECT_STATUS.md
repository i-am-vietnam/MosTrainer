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

## Verified This Inspection

- `ProjectValidator`: 21 projects, 0 errors, 0 warnings.
- Debug solution build: success, 0 reported errors.
- Current workbook lifecycle and Training flow were inspected from source.
- Testing Mode Phase 1 build and mode-selection logic: verified; Training remains the default and `Form1` is unchanged.
- Testing Mode Phase 2 selection, deduplication, insufficient-pool handling, fixed-order/index, deadline, remaining-time, expiration, and one-shot submission-state logic passed a disposable deterministic validation harness.
- No project JSON, starter workbook, assertion, or grading code was changed by Testing Mode Phase 1 or Phase 2.

## In Progress

- None.

## Planned

- Preserve Training behavior unchanged.
- Add Testing-only UI state, start the session after Testing login, and display project/countdown progress.
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

Phase 3+: **NOT IMPLEMENTED**

Testing login remains intentionally unwired and still shows the Phase 1 notice. No Testing UI, workbook navigation/persistence, submission, grading, or score calculation has been implemented.

## Next Recommended Task

After user review of Phase 2, implement Phase 3 only: Testing UI state, create the session from Login, display Project n/7, and show an MM:SS countdown derived from the session deadline. Do not add workbook persistence or real submission in that phase.
