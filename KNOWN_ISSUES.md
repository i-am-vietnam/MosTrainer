# Known Issues and Risks

Last reviewed: 2026-09-20.

## Confirmed Issues / Current Limitations

### Working workbook is reset when a project is opened

`Form1.OpenProjectExcel` always copies the starter over:

```text
Documents/MosTrainer/Working/<ProjectId>/work.xlsx
```

This is the existing Training behavior, but it cannot preserve answers when navigating among projects in a test.

### Application-driven workbook close does not save

`ExcelController.CloseWorkbook` calls `Workbook.Close(false)`. Unsaved workbook changes are discarded when Form1 closes Excel, including before the current project-open flow. Testing Mode needs an explicit save operation and error handling before navigation/submission.

### Main form cannot currently return to login

`LoginForm` hides itself, shows `Form1`, and closes itself when `Form1` closes. That ends the application. A completed test cannot currently return to a reusable Login screen.

### Authentication is an MVP stub

`LoginForm` accepts only hard-coded credentials `admin` / `123456`. There is no user repository or authentication service.

## Known Risks

### Grading is tied to one active Excel session

`ExcelController` owns one `ExcelSession`, and `GradingService` grades that live workbook. Final grading of seven projects must open/grade/close saved workbooks sequentially and must handle a failure without corrupting the remaining session results.

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
- `AppSession` stores only a static language string; lifecycle/reset semantics must be defined before adding mode/session state.
- `Form1._currentAssetsDir` records the deployed path but is not otherwise consumed by Form1.
- C# language version is not explicitly pinned in project files, although source comments target C# 7.3-compatible syntax.

## Questions / Needs Verification Before Testing Implementation

- Decide whether the user may navigate backward among the seven projects; the proposed model supports it, but the requirement explicitly mentions only Next Project.
- Decide whether closing MosTrainer mid-test abandons the attempt after confirmation or requires resume support.
- Decide whether Testing working directories are deleted immediately after result acknowledgement or retained temporarily for diagnostics/review.
- Decide how to present individual failed tasks after submission; the current requirement mandates total score, not detailed review.
- Verify save behavior and final grading with real Excel for a complete seven-project session after implementation.
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
