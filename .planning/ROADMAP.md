# Roadmap: BOHO Import Tool — Refactor BOM Logic

## Overview

Rebuild the BOM import logic for the BOHO Import Tool from a business-analysis-first, staging-safe foundation. Phase 1 locks down what to keep, fix, or discard from the current 11,133-line `main_window.py` based on real evidence from the 94 already-audited BOM sample files — no `core/` code is written until the business owner confirms these verdicts. Phase 2 then rebuilds the confirmed logic in `v3/core/` as pure, pytest-testable functions with a Staging-first SQL Server insert path, the four known bugs fixed, and full exception logging — all while the production `Tools/` app keeps running untouched.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

- [ ] **Phase 1: BA Analysis & Decision Lock** - Inventory every existing BOM logic path with real evidence and get user-confirmed keep/fix/discard verdicts before any code is written.
- [ ] **Phase 2: Core Rebuild — Architecture, Fixes & Observability** - Rebuild the confirmed BOM logic in `v3/core/` as staging-first, pytest-tested, logged pure functions with the four known bugs fixed.

## Phase Details

### Phase 1: BA Analysis & Decision Lock
**Goal**: Every existing BOM business-logic path is inventoried with real evidence and given an explicit, user-confirmed keep/fix/discard verdict — establishing the locked reference that Phase 2 implements against, before any `core/` code is written.
**Depends on**: Nothing (first phase)
**Requirements**: BA-01, BA-02, BA-03
**Success Criteria** (what must be TRUE):
  1. A full inventory of existing BOM logic (BTP detection, Fill-Forward, MKT fallback, SP_HOOK, section/header detection) exists, each item backed by real evidence from the 94-sample audit.
  2. Every inventoried logic item carries an explicit verdict: GIỮ NGUYÊN / SỬA (with reason + fix approach) / BỎ (with reason).
  3. The user, acting as business owner, has reviewed and confirmed all verdicts before any `core/` code is written.
  4. The confirmed verdicts are saved as a standalone reference document (not chat history) that Phase 2 can implement against.
**Plans**: 3 plans

Plans:
- [ ] 01-01-PLAN.md — Evidence base + tracer row G1: doc skeleton, changelog 2.2.1→2.2.25 (D-04), 94-sample + mapping metrics (D-02), live read-only SQL via mssql-boho (D-03)
- [ ] 01-02-PLAN.md — Full rule-level inventory (90+ rows, groups A-K) with evidence-backed GIỮ/SỬA/BỎ verdicts, §5 open questions, §6 Phase 2 mapping, integrity gate
- [ ] 01-03-PLAN.md — Business-owner review checkpoint + sign-off; 01-BA-DECISIONS.md locked as CONFIRMED

### Phase 2: Core Rebuild — Architecture, Fixes & Observability
**Goal**: The BOM logic confirmed in Phase 1 is rebuilt in `v3/core/` as testable, staging-first, fully-logged pure functions — with the four known bugs fixed — while `Tools/` production is never modified.
**Depends on**: Phase 1
**Requirements**: ARCH-01, ARCH-02, ARCH-03, ARCH-04, ARCH-05, FIX-01, FIX-02, FIX-03, FIX-04, OBS-01, OBS-02
**Success Criteria** (what must be TRUE):
  1. BOM business logic is implemented as pure functions in `core/`, runnable and testable via pytest with no CustomTkinter dependency.
  2. `ui/` (views) only calls into `core/` — it contains no BOM business logic itself.
  3. Inserting BOM data always flows Staging table → Validate → Transaction transfer (never a direct insert into Bravo business tables), and every raw Excel/DB value in the BOM flow passes through `core/sanitizer.py` (`safe_str`/`safe_int`/`safe_decimal`/`safe_date`).
  4. The four known bugs no longer reproduce against the 94 sample files: no `_mkt_cache` collisions, no garbage header columns, no `.strip()` crashes on numeric values, and optional sections (e.g. BOM5) no longer trigger a warning as if they were errors.
  5. Every exception in the new BOM flow is logged via `logger.exception()` to `logs/app.log` with enough detail (source file, section, row, error) to trace it, and the original `Tools/` production app remains unmodified throughout.
**Plans**: TBD

## Progress

**Execution Order:**
Phases execute in numeric order: 1 → 2

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. BA Analysis & Decision Lock | 0/3 | Planned | - |
| 2. Core Rebuild — Architecture, Fixes & Observability | 0/TBD | Not started | - |
