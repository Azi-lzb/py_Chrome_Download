# AGENTS.md

This file provides guidance to Codex when working in this repository.

## What this is

`ChromeAutoRun` is a single-file Selenium + openpyxl Windows automation tool for the HZPBC reporting workflow. The user runs `main.py` or a PyInstaller-built `ChromeAutoRun.exe`, then uses the interactive console menu to launch Chrome, log in, navigate to the report tree, and batch query/export reports.

## Commands

```powershell
# Run from source
python main.py

# Syntax check before packaging
& "$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe" -m py_compile main.py tests\test_matrix_recovery.py

# Unit test for the 3/4/5 recovery fallback helper
& "$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe" -m unittest tests.test_matrix_recovery -v

# Build exe with the HZPBC environment
& "$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe" -m PyInstaller --clean --noconfirm main.spec
```

## Packaging Rules

- **Always build with** `"$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe"`.
- Do not package with base conda `python`, because it may not have Selenium installed and can produce an exe that fails with `No module named selenium`.
- Before packaging, verify Selenium is available in the build environment:

```powershell
& "$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe" -c "import selenium; print(selenium.__version__)"
```

- `main.spec` must keep Selenium bundled. It imports `collect_submodules('selenium')`, includes those modules in `hiddenimports`, and intentionally fails the build if Selenium cannot be found.
- The PyInstaller output is `dist\ChromeAutoRun.exe`. The HZPBC deliverable is `dist_hzpbc\ChromeAutoRun.exe`.
- After a successful build, copy the exe to both:
  - `dist_hzpbc\ChromeAutoRun.exe`
  - `dist_hzpbc\ChromeAutoRun_YYYYMMDD_HHMMSS.exe`
- Keep old timestamped files in `dist_hzpbc` for rollback unless the user explicitly asks to clean them.

## Runtime Files

- `config.xlsx` and `config.json` live next to the running script/exe.
- Browser output and diagnostic files go under `output/`.
- `dist_hzpbc` is a deliverable folder and may contain its own `config.xlsx`.

## Important Menus

- `1` starts Chrome through `chromedriver.exe`.
- `2` logs in.
- `2.5` performs one-click launch + login + report-page navigation.
- `3`, `4`, and `5` run the three formal matrix export scenes with strong recovery fallback: `county`, `org`, and `county2`.
- Hidden aliases `3p`, `4p`, and `5p` run the old plain/no-recovery versions for `county`, `org`, and `county2`; do not show them in the menu.
- The old `t31`, `t41`, and `t51` enhanced recovery aliases are intentionally removed.
- The old low-level `t`, `t2`, and `t3` test menu entries are intentionally hidden for now; their helper functions may remain in code.

## Recovery Logic

The 3/4/5 export flow uses `_recover_report(scene, plan)` as the fallback chain:

1. soft refresh: tree navigation + iframe switch,
2. hard refresh: F5/reload + reestablish,
3. full recovery: login check/relogin + portal navigation + tree navigation + iframe switch.

Formal menu entries `3/4/5` use the recovery path. Hidden plain aliases `3p/4p/5p` bypass recovery and are kept only as a fallback for manual testing. When changing export behavior, keep `tests/test_matrix_recovery.py` passing.

## Editing Rules

- Keep all Python files UTF-8.
- Use `apply_patch` for manual code edits.
- Do not delete or overwrite user-created deliverables unless explicitly asked.
- Ignore unrelated dirty worktree files; do not revert them.
