# Management-PDO Review (Delta Hardening) — 2026-02-19

## Scope
Quick engineering hardening review focused on brittle AI-assisted fix risk, maintainability, and operational safety.

## Executive summary
The project has a solid baseline structure (routes/services split, path allowlisting, atomic Excel writes, SSE keepalive, CI smoke).
Main remaining risks are test coverage depth, auth boundary assumptions, and observability maturity.

## What was improved now

### 1) Extension policy centralized and tightened
**Symptom:** Extension checks were duplicated across routes and watcher, and included `.xls` despite write paths relying on `openpyxl`.

**Root cause hypothesis:** Validation logic was local per module, causing drift risk and format inconsistency.

**Invariant:** Supported Excel formats should be consistent across read/write/watch/open paths.

**Boundary ownership:** `app/config.py` (policy), consumed by routes/services.

**Change made:**
- Added `ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}` in `app/config.py`.
- Updated:
  - `app/routes/data.py`
  - `app/routes/excel.py`
  - `app/services/file_watcher.py`
- Error message now explicitly states supported extensions.

### 2) SSE client list concurrency hardening
**Symptom:** `connected_clients` list was mutated from async request context and watcher-driven reload paths without explicit client-list locking.

**Root cause hypothesis:** Shared mutable list had implicit concurrency assumptions.

**Invariant:** Client registration/removal/snapshot iteration must be race-safe.

**Boundary ownership:** `app/state.py` + SSE event/reload notification paths.

**Change made:**
- Added `clients_lock` in `app/state.py`.
- Guarded append/remove in `app/routes/events.py`.
- Guarded snapshot creation in `app/services/data_loader.py::_notify_clients`.

## Remaining prioritized findings

### High
1. **No automated regression suite beyond compile/import smoke**
   - Risk: brittle fixes can pass CI and fail behaviorally.
   - Recommendation: add API regression tests for `/api/save-task`, `/api/add-task`, `/events`, and path-guard checks.

2. **No auth/authorization around mutating endpoints**
   - Risk: safe for localhost-only usage; unsafe if exposed to broader network.
   - Recommendation: enforce loopback binding + optional API token/mTLS/reverse-proxy auth.

### Medium
3. **Global process state scales poorly** (`cached_data`, `data_version`, client registry)
   - Recommendation: encapsulate state behind a small state service object and explicit synchronization policy.

4. **Observability mostly print-based**
   - Recommendation: migrate to structured logging with request/file/sheet/version correlation keys.

5. **No explicit rate/size limits on update payloads**
   - Recommendation: enforce max field sizes and max new columns per request.

### Low
6. **Documentation still implies broader Excel support in places**
   - Recommendation: align docs fully with `.xlsx/.xlsm` write-safe support policy.

## Suggested next sprint hardening plan
1. Add pytest + FastAPI TestClient suite (regression + negative path).
2. Add at least one generalized test matrix (invalid columns, stale row index conflicts, concurrent save simulation).
3. Add optional authentication layer for mutating routes.
4. Add structured log events around save/add/reload and SSE notify.
5. Refactor global state into a service module with clear ownership.

## Validation performed
- Python compile check passed (`python -m compileall -q .`).
- Manual code inspection of config, routes, watcher, loader, SSE paths.
