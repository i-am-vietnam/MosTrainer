# Known Issues and Risks

Last reviewed: 2026-09-21.

## Confirmed Issues / Current Limitations

### Training working workbook is reset when a project is opened

`Form1.OpenProjectExcel` always copies the starter over:

```text
Documents/MosTrainer/Working/<ProjectId>/work.xlsx
```

This remains the existing Training behavior. Testing no longer uses this path: Phase 4 uses a session-scoped, copy-once working file and reopens it without overwriting learner work.

### Base Excel close operation does not save

`ExcelController.CloseWorkbook` still calls `Workbook.Close(false)` to preserve Training behavior. Testing Phase 4 explicitly calls `SaveWorkbook` before every navigation/form close and before timeout close; this ordering must remain mandatory in submission code.

### Timeout submission failure has no retry UI

Timeout now invokes the shared submission pipeline automatically. If saving, preparing, opening, or grading infrastructure fails after the deadline, the attempt is aborted without a score and the exam remains locked at `00:00`. There is intentionally no automatic retry loop or retry button yet.

### Failed, abandoned, or crash-interrupted Testing workspaces may remain

Successful manual and timeout submissions delete only their current session directory after the learner acknowledges the result. Failed submissions, confirmed mid-test exits, and crash-interrupted sessions are intentionally retained for diagnostics. There is no automatic stale-workspace sweep or resume support.

### Authentication is an MVP stub

`LoginForm` accepts only hard-coded credentials `admin` / `123456`. There is no user repository or authentication service.

## Known Risks

### Grading is tied to one active Excel session

`ExcelController` owns one `ExcelSession`, and `GradingService` grades that live workbook. Manual and timeout submission both open/grade/close the seven saved workbooks sequentially through the same pipeline; any future retry/recovery UI must preserve that ownership model.

### Excel Interop lifecycle is failure-sensitive

The application owns a visible Excel process and may force-terminate its recorded PID after COM cleanup. Manual Excel closure, COM disconnection, save failure, file locks, or a grading exception need explicit Testing-mode recovery and user messaging.

Controlled Phase 7 checks confirmed that manually closing the owned Excel window before Next Project keeps the old project index and shows a clear switch error. Doing so before Submit aborts submission, creates no score, and retains the workspace. The disconnected COM session is not automatically reattached; the learner must exit the test, so this remains a recoverability limitation rather than a crash or silent data-loss path.

### Testing session state is not persistent

The Phase 2 `TestSession` is in memory only. If the learner closes MosTrainer or the process crashes mid-test, the session is lost; resume is not implemented. Decide whether the first release explicitly abandons the test or a later phase adds a session manifest/resume feature.

### Project validity may change after session creation

`ProjectLoader` returns only valid packages at load time. A package/starter/language file could still become unavailable after random selection. Testing startup and every open/save/submit transition need guarded failure handling.

Testing Restart also depends on the selected project's starter remaining available. A missing starter or failed replacement reports an error and attempts to reopen the existing working file; if Excel cannot reopen it, the Testing UI remains locked while the deadline continues. There is no dedicated reopen/retry UI.

### UI localization is partial

Task text is localized through project language JSON, and Login labels switch language, but much of Form1 status/chrome is hard-coded English. New Testing UI must at least provide EN/VI text without changing project task JSON.

## Technical Debt

- `MosTrainer.Data` contains SQLite initialization/result persistence, but the current UI does not initialize or use it.
- `Form1` contains UI and workflow orchestration in one class; Testing should add small session/submission helpers rather than duplicate Form1 or perform a broad refactor.
- `ExcelController.cs` is very large and COM-sensitive. Do not split/refactor it as part of Testing Mode.
- `AppSession` stores static language/mode state, while each Testing login creates and passes a fresh TestSession directly to Form1.
- `Form1._currentAssetsDir` records the deployed path but is not otherwise consumed by Form1.
- C# language version is not explicitly pinned in project files, although source comments target C# 7.3-compatible syntax.

## Questions / Future Product Decisions

- Supervised Training Grade popup verification remains open: the disposable real-Excel harness stalled inside a grading call before it could observe the EN/VI popup. Source review confirms one `CheckTask` call and the localized popup mapping, but it is not an end-to-end PASS yet.
- The final result exposes only failed task identifiers, not expected answers, assertion messages, or detailed review guidance. Decide whether a future instructor-only review screen is needed.
- Decide whether retained failed/abandoned workspaces need an administrative cleanup tool or a resume/session-manifest feature.

## Edge Cases Required in the Testing Plan

- Fewer than seven valid projects or duplicate ProjectIds.
- A valid package with zero tasks.
- Missing/invalid metadata, starter, task, language, or optional asset files.
- Excel open/save/close failure, locked workbook, or manually closed Excel.
- Application exit during a test.
- Project switch failure after current workbook save.
- One task or whole-project grading exception.
- Repeated Submit clicks or navigation during submission.
- Score rounding and the all-pass exactly-1000 invariant.
- Reusing an ended session or starting a second test after return to Login.
- `Training -> Testing -> Training` in one application lifetime.
