# Current Status

Last verified: 2026-09-21
Branch: `master`
Commit: `8a4d1df4c57d7f9dfee7bd341e31067deeae89af` (Phase 5 baseline)

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
- Testing Mode Phase 4 is complete: each session owns isolated, copy-once project workbooks; Testing opens real Excel/task instructions and safely saves, closes, and reopens them through Previous/Next Project navigation.
- Testing Mode Phase 5 is complete: manual Submit uses one shared seven-workbook grading pipeline, collects task/project results, calculates the equal-project-weight score out of 1000, shows a bilingual result, and returns to the existing LoginForm.
- Testing Mode Phase 6 is complete: the 50-minute deadline automatically invokes the same submission pipeline with `TimeExpired`, without confirmation.

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
- No project JSON, starter workbook, assertion, or grading code was changed by Testing Mode Phase 1 through Phase 6.

## In Progress

- None.

## Planned

- Preserve Training behavior unchanged.
- Add completed-session workspace cleanup when its retention policy is defined.
- Perform final UI polish and full end-to-end Training/Testing regression.

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
- Each starter is copied only on the project's first visit; revisits open the existing working file without overwriting it.
- Testing opens real Excel, deploys existing project assets, and displays localized task tabs with existing Previous/Next Task navigation.
- Separate Previous/Next Project controls perform explicit Save -> Close -> Open; the session index and UI commit only after the target workbook opens successfully.
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

Phase 7+: **NOT IMPLEMENTED**

Completed Testing workspace cleanup/retention policy, final UI polish, and final broad acceptance regression remain deferred.

## Next Recommended Task

After user review of Phase 6, define the completed-workspace retention/cleanup policy, then perform final UI polish and broad Training/Testing acceptance regression without changing grading semantics.
