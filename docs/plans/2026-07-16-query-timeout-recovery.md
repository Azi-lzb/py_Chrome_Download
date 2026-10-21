# Query Timeout Recovery Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Treat a report query as stalled only when it remains in the loading state beyond that scene's `task.query_wait_seconds`, then recover the report and retry the same matrix item.

**Architecture:** `_fr_wait_query_done` already reports both elapsed time and whether loading was observed. Add a small pure helper to distinguish an active-loading timeout from a completed-page timeout, then let the recovery-mode matrix loop use that helper before export. The configured query wait remains the sole query timeout; no global timeout is introduced.

**Tech Stack:** Python 3.11, Selenium, unittest, openpyxl.

---

### Task 1: Define stalled-query classification

**Files:**
- Modify: `tests/test_matrix_recovery.py`
- Modify: `main.py`

**Step 1: Write the failing test**

Add tests that classify `{"timeout": true, "saw_loading": true}` as a stalled query and `{"timeout": true, "saw_loading": false}` as non-stalled.

**Step 2: Run test to verify it fails**

Run: `& "$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe" -m unittest tests.test_matrix_recovery -v`

Expected: FAIL because the new classification helper does not exist.

**Step 3: Write minimal implementation**

Add a pure `_query_wait_indicates_stall(wait_result)` helper that returns true only for a timeout with loading observed at the timeout boundary.

**Step 4: Run test to verify it passes**

Run the same unittest command. Expected: PASS.

### Task 2: Recover and retry the current item

**Files:**
- Modify: `main.py`

**Step 1: Apply the classification after query waiting**

When recovery mode is enabled and the query is classified stalled, append a query-stalled log record, call `_recover_report(scene, plan, validate_label=...)`, and rerun the same matrix coordinate instead of exporting or advancing to the next item.

**Step 2: Preserve non-stalled timeout behavior**

If the configured wait elapsed but loading is no longer active, retain the existing log and normal export path.

**Step 3: Verify syntax and regression tests**

Run:

```powershell
& "$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe" -m py_compile main.py tests\test_matrix_recovery.py
& "$env:USERPROFILE\miniconda3\envs\HZPBC\python.exe" -m unittest tests.test_matrix_recovery -v
```

### Task 3: Deliverable verification

**Files:**
- Modify: `main.py`
- Modify: `tests/test_matrix_recovery.py`

**Step 1: Review the diff**

Confirm the new behavior uses only `task.query_wait_seconds` and only applies recovery to `recovery_mode='recover'`.

**Step 2: Package only on request**

Run the HZPBC PyInstaller build and copy to `dist_hzpbc` only if a new executable is requested.
