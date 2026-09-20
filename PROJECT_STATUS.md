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

## Verified This Inspection

- `ProjectValidator`: 21 projects, 0 errors, 0 warnings.
- Debug solution build: success, 0 reported errors.
- Current workbook lifecycle and Training flow were inspected from source.
- Testing Mode Phase 1 build and mode-selection logic: verified; Training remains the default and `Form1` is unchanged.
- No project JSON, starter workbook, assertion, or grading code was changed by Testing Mode Phase 1.

## In Progress

- None.

## Planned

- Preserve Training behavior unchanged.
- Add an in-memory seven-project test session with fixed unique randomized projects.
- Add non-destructive, session-scoped workbook persistence and project navigation.
- Disable per-task grading feedback in Testing.
- Add final submission, reuse of current grading, 1000-point calculation, result display, cleanup, and return to login.

## Testing Mode

Phase 1: **COMPLETED / VERIFIED**

Phase 2+: **NOT IMPLEMENTED**

Phase 1 adds `AppMode`, `AppSession.Mode`, and Login mode selection. Selecting Testing stores the mode, shows a clear next-phase notice, and does not enter the Training workflow. No Testing session, project randomization, navigation, persistence, submission, grading, or score calculation has been implemented.

## Next Recommended Task

After user review of Phase 1, implement Phase 2 only: pure Testing session/project/result models and fixed seven-project selection tests, without workbook persistence or submission.
