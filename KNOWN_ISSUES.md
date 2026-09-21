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

### Completed Testing workspaces are retained

After a successful manual submission, `Documents/MosTrainer/Testing/<SessionId>` is not automatically deleted. This is an intentional deferred cleanup/retention decision, not a grading failure.

### Authentication is an MVP stub

`LoginForm` accepts only hard-coded credentials `admin` / `123456`. There is no user repository or authentication service.

## Known Risks

### Grading is tied to one active Excel session

`ExcelController` owns one `ExcelSession`, and `GradingService` grades that live workbook. Manual and timeout submission both open/grade/close the seven saved workbooks sequentially through the same pipeline; any future retry/recovery UI must preserve that ownership model.

### Excel Interop lifecycle is failure-sensitive

The application owns a visible Excel process and may force-terminate its recorded PID after COM cleanup. Manual Excel closure, COM disconnection, save failure, file locks, or a grading exception need explicit Testing-mode recovery and user messaging.

### Testing session state is not persistent

The Phase 2 `TestSession` is in memory only. If the learner closes MosTrainer or the process crashes mid-test, the session is lost; resume is not implemented. Decide whether the first release explicitly abandons the test or a later phase adds a session manifest/resume feature.

### Project validity may change after session creation

`ProjectLoader` returns only valid packages at load time. A package/starter/language file could still become unavailable after random selection. Testing startup and every open/save/submit transition need guarded failure handling.

### UI localization is partial

Task text is localized through project language JSON, and Login labels switch language, but much of Form1 status/chrome is hard-coded English. New Testing UI must at least provide EN/VI text without changing project task JSON.

## Technical Debt

- `MosTrainer.Data` contains SQLite initialization/result persistence, but the current UI does not initialize or use it.
- `Form1` contains UI and workflow orchestration in one class; Testing should add small session/submission helpers rather than duplicate Form1 or perform a broad refactor.
- `ExcelController.cs` is very large and COM-sensitive. Do not split/refactor it as part of Testing Mode.
- `AppSession` stores static language/mode state, while TestSession is passed directly to Form1. Final lifecycle/reset semantics still need to be defined for completed or abandoned tests.
- `Form1._currentAssetsDir` records the deployed path but is not otherwise consumed by Form1.
- C# language version is not explicitly pinned in project files, although source comments target C# 7.3-compatible syntax.

## Questions / Needs Verification Before Testing Implementation

- Decide whether closing MosTrainer mid-test abandons the attempt after confirmation or requires resume support.
- Decide whether Testing working directories are deleted immediately after result acknowledgement or retained temporarily for diagnostics/review.
- Decide whether a future review screen should expose individual failed tasks; Phase 5 intentionally displays only total score.
- Manual and timeout seven-project grading/submission have been verified with real Excel. Final broad end-to-end acceptance across all lifecycle combinations remains planned.
- Verify behavior when the learner manually closes the owned Excel workbook/application, then attempts Next or Submit.

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
