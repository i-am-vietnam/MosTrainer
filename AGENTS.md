# MosTrainer Agent Guide

This file applies to the entire repository. Keep it short, enforceable, and current.

## Start Here

At the beginning of a new conversation, read in this order:

1. `AGENTS.md`
2. `PROJECT_STATUS.md`
3. `KNOWN_ISSUES.md`
4. `CODEX_CONTEXT.md`
5. Only the source files directly relevant to the current task

Do not rescan the whole repository unless the task or conflicting evidence makes that necessary.

## Source of Truth

Use this precedence order:

```text
actual source code
  > PROJECT_STATUS.md
  > CODEX_CONTEXT.md
  > conversation history
```

If documentation conflicts with source, source wins. Correct the affected documentation as part of the scoped work.

## Protected Subsystem: Grading

Treat the following as stable and protected:

- `MosTrainer.Core/Services/GradingService.cs`
- `MosTrainer.Core/Interfaces/IExcelController.cs` assertion surface
- existing assertion implementations in `MosTrainer.Excel/ExcelController.cs`
- existing PASS/FAIL behavior
- validated `Projects/Excel2019_Pxx/tasks.json` definitions
- `MosTrainer.Projects/ProjectValidator.cs` assertion validation

Do not modify grading, assertions, project task definitions, or established PASS/FAIL semantics unless the user explicitly requests a grading change. If a grading change appears necessary:

1. identify the exact file and method;
2. explain why configuration or existing behavior is insufficient;
3. state regression risk and the old projects/tasks affected;
4. propose focused positive and negative regression tests;
5. wait for user approval before editing.

New application modes must orchestrate and reuse `GradingService.CheckTask`; they must not duplicate or replace the grading system.

## Change Discipline

Prefer minimal change and low regression risk. Do not perform unrelated cleanup, mass formatting, broad renames, namespace moves, framework changes, technology changes, dependency upgrades, or rewrites of working code.

Preserve these baselines unless the user explicitly changes them:

- C# WinForms on .NET Framework 4.7.2
- Microsoft Excel Interop
- folder-based project packages and their JSON schema
- current Training behavior
- existing project starter workbooks

Never edit a `starter.xlsx` in place. Use a copy for experiments.

## Before Coding a Significant Feature

1. Read the four context files above.
2. Inspect only the relevant source and current git status.
3. Describe the current flow with file/class/method names.
4. List files expected to change.
5. List protected files that must not change.
6. Write a small implementation and verification plan.
7. Then code, if the user authorized implementation.

Do not infer that a planning request authorizes implementation.

## After a Significant Change

- Build `MosTrainer/MosTrainer.slnx`.
- Run focused positive, negative, persistence, and regression checks proportional to the change.
- Run `ProjectValidator` when project packages or assertion routing are affected.
- Review `git diff --check`, `git status`, and the complete scoped diff.
- Update `PROJECT_STATUS.md`.
- Update `CODEX_CONTEXT.md` if architecture changed.
- Update `KNOWN_ISSUES.md` only for confirmed issues, concrete risks, debt, or explicit verification questions.

Change `AGENTS.md` only when project-wide development rules change. Do not turn `CODEX_CONTEXT.md` into a commit log, `PROJECT_STATUS.md` into a long architecture document, or `KNOWN_ISSUES.md` into speculation.

## Safety and Git

- Preserve user changes in a dirty worktree.
- Do not reset, restore, checkout over, commit, or push unless explicitly asked.
- Do not leave test workbooks, temporary extraction directories, Excel locks, or orphan `EXCEL.EXE` processes.
- Use paths relative to the repository root in documentation.

