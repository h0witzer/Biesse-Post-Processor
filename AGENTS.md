# AGENT INSTRUCTIONS

## Scope
These instructions apply to the entire repository.

## Code style and structure
- Keep each VBA module or form using `Option Explicit` and avoid `On Error Resume Next`; handle errors deliberately and close resources (files, selections, dialog state) explicitly.
- Preserve existing naming and metadata conventions (e.g., release history, customer info, and constants in `Post.bas`). When adding routines, prefer small, single-purpose procedures with clear comments about their Alphacam context.
- When touching UI forms, keep control names stable and update any code-behind that references them.

## Testing and tooling
- Use the Python-based tooling in `requirements-dev.txt` for local checks (e.g., macro parsing with `oletools` or future pytest suites). If you add tests, run `python -m pytest` from the repo root.

## PR / review expectations
- Keep PR messages concise: a short summary of functional changes plus a bullet list of tests or checks you ran.
- In commit messages, describe the intent of the change, not just the files edited.
