# Current Status

Last verified: 2026-09-21
Branch: `master`
Commit/baseline: `fbde3b207143549679ffcaaf76690e2969dd24d8` (Phase 7 HEAD; Phase 8 changes are local and uncommitted)

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
- Testing Mode Phase 4 is complete: each session owns isolated, copy-once project workbooks; Testing opens real Excel/task instructions and safely saves, closes, and opens the next project workbook.
- Testing Mode Phase 5 is complete: manual Submit uses one shared seven-workbook grading pipeline, collects task/project results, calculates the equal-project-weight score out of 1000, shows a bilingual result, and returns to the existing LoginForm.
- Testing Mode Phase 6 is complete: the 50-minute deadline automatically invokes the same submission pipeline with `TimeExpired`, without confirmation.
- Testing Mode Phase 7 is complete: successful completed-session workspaces are cleaned only after result acknowledgement; failed, abandoned, and crash-interrupted workspaces are retained; mid-test close is confirmed; and the Testing footer is bilingual and explicit.
- Testing Mode Phase 8 is complete: the exam countdown is centered above task instructions, project navigation is strictly forward-only, and the final result lists failed task identifiers without exposing diagnostic details.

## Verified This Inspection

- `ProjectValidator`: 21 projects, 0 errors, 0 warnings.
- Debug solution build: success, 0 reported errors.
- Current workbook lifecycle and Training flow were inspected from source.
- Testing Mode Phase 1 build and mode-selection logic: verified; Training remains the default and `Form1` is unchanged.
- Testing Mode Phase 2 selection, deduplication, insufficient-pool handling, fixed-order/index, deadline, remaining-time, expiration, and one-shot submission-state logic passed a disposable deterministic validation harness.
- Testing Mode Phase 3 Login-to-Form integration, EN/VI UI, project progress, countdown, timeout lock, Training constructor path, and absence of Excel startup passed a disposable WinForms integration harness.
- Testing Mode Phase 4 passed a disposable real-Excel integration harness covering Login/session transfer, initial workbook/task tabs, `1 -> 2 -> 3 -> 2 -> 1`, persisted workbook content, boundary button states, session-path isolation, unchanged session timer identity, timeout save/close/lock, untouched source starter, and no new orphan Excel process.
- Testing Mode Phase 5 passed pure score/state checks plus disposable real-Excel submission checks: confirmation No, seven projects/52 tasks graded exactly once, five unvisited projects initialized, saved current working content, result OK returning to Login, failure rollback/retry, unchanged starters, and no orphan Excel process.
- Testing Mode Phase 6 passed disposable real-Excel timeout checks at Projects 1, 4, and 7; each graded seven projects and the exact selected task total, initialized unvisited workbooks, returned a `TimeExpired` result, returned to Login, and left no orphan Excel process. Timeout failure, repeated timeout, manual No/Yes, and new-session-after-Login checks also passed.
- Testing Mode Phase 7 passed disposable real-Excel checks for manual and timeout completion/cleanup, cleanup isolation/idempotence, retained failure/abandon workspaces, EN/VI close confirmation and controls, saved-workbook persistence, manual Excel-close error paths, score invariants, unchanged starter hashes, and no remaining owned Excel process.
- Testing Mode Phase 8 passed disposable real-Excel checks for centralized EN/VI countdown display, forward-only `1 -> 2 -> ... -> 7`, backward/skip rejection, preserved earlier workbooks, Project 7 button state, manual/timeout result formatting, post-result cleanup, and no orphan Excel process. Synthetic checks covered perfect results, identifier fallback, stable order, and six-identifiers-per-line wrapping.
- Training EN/VI shell plus disposable Go/task navigation/Grade/Restart/elapsed-timer/close behavior passed without touching an existing Training working directory.
- No project JSON, starter workbook, assertion, grading code, submission grading behavior, or score formula was changed by Testing Mode Phase 1 through Phase 8.

## In Progress

- None.

## Planned

- Classroom acceptance testing on the target deployment image and representative learner machines.

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

Phase 4: **COMPLETED / VERIFIED**

- Testing workbooks use `Documents/MosTrainer/Testing/<SessionId>/<ProjectId>/work.xlsx`.
- Each starter is copied only when its project is first reached; existing working files are never overwritten.
- Testing opens real Excel, deploys existing project assets, and displays localized task tabs with existing Previous/Next Task navigation.
- Project switching performs explicit Save -> Close -> Open; the session index and UI commit only after the next workbook opens successfully.
- Timeout locks Testing, attempts to save, and closes Excel only after a successful save. It does not start submission or grading.

Phase 5: **COMPLETED / VERIFIED**

- Submit is available only on Project 7/7 and requires bilingual Yes/No confirmation.
- Manual submission saves/closes the active workbook, initializes untouched workbooks for unvisited projects, and grades all seven in fixed order through the existing `GradingService.CheckTask`.
- Task/project/submission result models retain diagnostic results internally without showing individual PASS/FAIL to the learner.
- Score uses equal project weights and decimal arithmetic; perfect is exactly 1000, while a non-perfect displayed score is capped below 1000.
- Successful result acknowledgement closes Testing Form1 and shows the existing LoginForm. Infrastructure failure aborts submission and restores the workbook/timer for retry when time remains.

Phase 6: **COMPLETED / VERIFIED**

- At `00:00`, Testing sets one-shot timeout state, stops the timer, locks interactions, and calls `BeginTestingSubmission(TestSubmissionReason.TimeExpired)` directly.
- Timeout shows no confirmation and reuses the exact manual Save/Close, seven-workbook grading, score, result, and Login-return pipeline.
- Timeout works from any current project; unvisited projects are initialized as untouched session workbooks and graded normally.
- Infrastructure failure after deadline aborts submission without a fake score, keeps `00:00`, and leaves all Testing interactions locked.

Phase 7: **COMPLETED / VERIFIED**

- Result acknowledgement is the cleanup boundary for successful manual and timeout sessions.
- Cleanup is path-checked, idempotent, and deletes only the current SessionId directory; other Testing sessions and the Testing root remain intact.
- Failed, abandoned, and crash-interrupted workspaces remain available for diagnostics. No global stale-directory sweep exists.
- Closing an incomplete test asks for bilingual confirmation. No resumes the same deadline/workbook; Yes best-effort saves, closes owned Excel, returns to Login, and retains the abandoned workspace.
- Testing task/project navigation labels are explicit in EN/VI and Submit has a distinct accent without changing Training control text/behavior.

Phase 8: **COMPLETED / VERIFIED**

- Testing uses a dedicated centered `lblTestingCountdown` row immediately above task tabs; the Training footer timer remains separate and the new row is collapsed in Training.
- Testing project navigation is forward-only. The Previous Project control is hidden/disabled, and the backend accepts only `targetIndex == CurrentProjectIndex + 1`; backward and skipped targets are rejected without changing workbook or session index.
- Earlier project working files remain in the session workspace for final grading after the learner advances.
- Manual and timeout results display total score plus failed identifiers such as `P01-Task2`, preserving project/task order and wrapping six identifiers per line. Diagnostic assertion messages are never displayed.

## Next Recommended Task

Run supervised classroom acceptance testing on the intended Windows/Office deployment image. Do not change grading semantics unless a separately approved grading defect is reproduced.
