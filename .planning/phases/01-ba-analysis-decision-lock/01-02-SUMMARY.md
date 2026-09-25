---
phase: 01-ba-analysis-decision-lock
plan: 02
subsystem: documentation
tags: [bom, mssql-server, bravo-erp, evidence-audit, decision-document, bom-parser, ast-sweep]

# Dependency graph
requires:
  - phase: 01-ba-analysis-decision-lock/01-01
    provides: "01-BA-DECISIONS.md §0-§3 (evidence base: EV-S00..EV-S14, EV-M01, SQL-00..SQL-08) and the tracer row G1"
provides:
  - "01-BA-DECISIONS.md §4 complete: 90 rule-level GIỮ/SỬA/BỎ verdict rows (A1-A7, B1-B5, C1-C10, D1-D4/D5a-g/D6-D8, E1-E14, F1-F12, G1-G5, H1-H8, I1-I4, J1-J6, K1-K5), each with code citation + evidence ID"
  - "§5 open questions consolidated: Q-01..Q-07 (Q-01..Q-04 carried from Plan 01, Q-05/Q-06/Q-07 added), every question has a Trả lời slot for Plan 03; explicit EV-M-only SỬA/BỎ list (D2, D3, D7, E5)"
  - "§6 summary + Phase 2 mapping: GIỮ 71 · SỬA 16 · BỎ 3 = 90; SỬA/BỎ → requirement mapping table; FIX-01..FIX-04 all confirmed present; 9 Cao-risk rows listed"
  - "§8 appendix: K1 (.strip() sweep, 69 unguarded/41 guarded locations) and K2 (except-without-log sweep, 43 locations), file:line:func:snippet, generated from a read-only AST script (deleted after use)"
  - "Full automated integrity-gate verify (embedded in 01-02-PLAN.md Task 3) passes"
affects: [01-03-business-owner-signoff, phase-02-core-refactor]

# Actuals (#2632)
actuals:
  tokens: 18500
  tasks: 3
  commits: 3

tech-stack:
  added: []
  patterns:
    - "AST-based read-only sweep script (ast.walk over named function spans) for cataloging risk patterns (.strip() receiver safety, except-without-log) across specific functions in a large legacy file, instead of manual grep — matches CLAUDE.md's 'heavy sweep → temp script → read report once' rule"
    - "Call-site enumeration via grep as evidence for 'this code path is dead for BOM' claims (D2/D3, B5) — a structural cross-reference distinct from EV-Snn sample evidence and SQL-nn live-data evidence, cited inline per §0.2.e's 'có thể thêm...' allowance"

key-files:
  created:
    - .planning/phases/01-ba-analysis-decision-lock/01-02-SUMMARY.md
  modified:
    - .planning/phases/01-ba-analysis-decision-lock/01-BA-DECISIONS.md

key-decisions:
  - "Corrected a planning-time citation error inherited from the plan spec: §4.C (row parse) was specified against `_parse_section_excel_rows` (bom_parser.py:297-407), which 01-01-SUMMARY.md/EV-S13/EV-S08 already established is THDM-only. Rewrote C1-C10 against the real BOM function `_parse_sheet` (bom_parser.py:461-588) and flagged the discrepancy inline per affected row, per the plan's own 're-read the code and correct the line references' instruction."
  - "Traced `_resolve_row_mapping`'s only 2 call sites in the repo (grep) to confirm its `Excel`/`Bien_doi` branches (D2/D3) never execute for BOM — BOM's `_resolve_detail_row` call passes a pre-filtered `_non_excel` list and `get_excel_val=None`; only THDM's `_resolve_thdm_thvt_row` supplies `get_excel_val`. Verdict BỎ for D2/D3 (don't port), not a rubber-stamped GIỮ — BOM's real Excel-macro handling is `_resolve_detail_row` Pass 1b (E5/E6/E7)."
  - "Corrected Plan 01's own SQL-06 evidentiary note (§3): its interpretation speculated 'no transaction wraps header+detail insert' based on 8 historical orphaned headers. Reading `_run_insert_bg` (main_window.py:10839-10874) directly shows the current code already wraps header insert + `_generate_bom_details` + a single `conn.commit()` in one transaction. K3's verdict stays SỬA, but for the real remaining gap (ARCH-03 staging-table validation, not transaction wrapping)."
  - "K1/K2 (69 unguarded `.strip()` sites, 43 silent `except` blocks) were built from a temporary read-only AST script (`tests/tmp/ba_bom_flow_sweep.py`, deleted per Task 3 Step 6) scanning exactly the BOM-flow function spans the plan named — not from manual reading, per CLAUDE.md's heavy-sweep rule."
  - "Did not mark BA-02 complete in requirements-completed, despite the plan frontmatter listing it. BA-02 requires the document's own §7 business-owner sign-off, explicitly deferred to Plan 03 ('Im lặng, auto-advance hay yolo mode không tính là xác nhận'). Marking it complete here would be the executor asserting a confirmation no human has given."

requirements-completed: []

coverage:
  - id: D1
    description: "§4.A-§4.D (36 rows): sheet/section detection, header detection, row parse (rewritten against the real _parse_sheet function), NguonDL resolve — every row evidence-backed, all committed"
    requirement: BA-01
    verification:
      - kind: other
        ref: "Task 1 <verify> automated script embedded in 01-02-PLAN.md (36-row presence, verdict/risk token, code-ref + EV/SQL citation assertions)"
        status: pass
    human_judgment: false
  - id: D2
    description: "§4.E-§4.G/§4.I (35 rows): _resolve_detail_row, BTP detection, MKT fallback G2-G5 (G1 kept from Plan 01), insert-time Fill-Forward"
    requirement: BA-01
    verification:
      - kind: other
        ref: "Task 2 <verify> automated script embedded in 01-02-PLAN.md (35-row presence + F7-F11 EV-S09 citation assertions)"
        status: pass
    human_judgment: false
  - id: D3
    description: "§4.H/§4.J/§4.K (19 rows) + §5 consolidation + §6 summary/mapping + §8 sweep appendix + full 90-row integrity gate"
    requirement: BA-01
    verification:
      - kind: other
        ref: "Task 3 <verify> automated integrity-gate script embedded in 01-02-PLAN.md (all 90 rows, counts-match-§6, D-04 chain citation, Q-answer-slot, FIX-01..04 presence, temp-file-deleted assertions)"
        status: pass
    human_judgment: false
  - id: D4
    description: "Document is ready for business-owner review (BA-02) — still DRAFT, sign-off explicitly deferred to Plan 03"
    requirement: BA-02
    verification: []
    human_judgment: true
    rationale: "BA-02 requires an actual business-owner confirmation recorded in §7 with a date, per the document's own convention. No such confirmation exists yet — Plan 03's job, not something this plan can self-certify."

duration: 25min
completed: 2026-09-25
status: complete
---

# Phase 1 Plan 02: Rule-Level BOM Logic Inventory Summary

**Completed the 90-row rule-level GIỮ/SỬA/BỎ inventory (§4.A-§4.K) of every BOM parsing/resolve/insert branch, each verdict backed by the 94-sample/SQL/git evidence base Plan 01 built — including 2 corrections to Plan 01's own evidentiary assumptions (C-group's real parse function, SQL-06's transaction-wrapping claim) discovered while re-reading the code as instructed.**

## Performance

- **Duration:** 25 min (2026-09-25T06:39:33Z → 2026-09-25T07:04:54Z)
- **Started:** 2026-09-25T06:39:33Z
- **Completed:** 2026-09-25T07:04:54Z
- **Tasks:** 3/3
- **Files modified:** 1 (`01-BA-DECISIONS.md`)

## Accomplishments

- **90 rule-level verdict rows written** (A1-A7, B1-B5, C1-C10, D1-D4/D5a-g/D6-D8, E1-E14, F1-F12, G1-G5, H1-H8, I1-I4, J1-J6, K1-K5) — exactly the mandatory minimum, no rubber-stamped extras. Final tally: **GIỮ 71 · SỬA 16 · BỎ 3**. 9 rows carry Mức rủi ro = Cao (`A4`, `C2`, `G1`, `G2`, `G3`, `G4`, `J4`, `K1`, `K2`).
- **Corrected 2 evidentiary assumptions inherited from planning-time citations**, both caught by re-reading the actual code rather than trusting the plan's line references (as the plan itself instructed): (1) §4.C's row-parse group was specified against a THDM-only function; rewrote against BOM's real `_parse_sheet`. (2) SQL-06's own note in §3 (Plan 01) speculated no transaction wraps header+detail insert; direct code reading of `_run_insert_bg` shows it already does — K3's fix target moved to the real gap (ARCH-03 staging, not transaction wrapping).
- **Traced dead-code-for-BOM branches via call-site enumeration** (not sample data): `_resolve_row_mapping`'s `Excel`/`Bien_doi` branches (D2/D3) and the `MucLookup` branch (D7) never execute for BOM's actual call path — verdict BỎ, explicitly not ported into `core/`, with the real BOM-equivalent mechanism cross-referenced (E5/E6/E7).
- **K1/K2 built from a read-only AST sweep script** (`tests/tmp/ba_bom_flow_sweep.py`, per CLAUDE.md's heavy-sweep rule) over the exact BOM-flow function spans named in the plan: 69 unguarded `.strip()` call sites (the exact bug class `8a07639` crashed on once) and 43 silent `except` blocks, cataloged file:line:function in §8.1/§8.2. Script and its JSON output deleted after the appendix was written.
- **§5/§6 consolidated**: added `Q-05` (A4 meta-key source priority), `Q-06` (A5 optional-section list), `Q-07` (I4 Fill-Forward-as-category confirmation) to the 4 questions Plan 01 recorded; explicitly listed the 4 SỬA/BỎ rows whose only evidence is `EV-M01` (D2, D3, D7, E5) per the evidence rule; wrote the full SỬA/BỎ → Phase 2 requirement mapping table with FIX-01..FIX-04 all confirmed present.

## Task Commits

Each task was committed atomically:

1. **Task 1: §4.A-§4.D parser/resolver inventory (36 rows)** - `5ba67ea` (feat)
2. **Task 2: §4.E-§4.G/§4.I insert-time BOM logic (35 rows)** - `aafa893` (feat)
3. **Task 3: §4.H/§4.J/§4.K + §5-§6/§8 consolidation + integrity gate (19 rows)** - `de5e2fc` (feat)

**Plan metadata:** commit pending (this SUMMARY + STATE/ROADMAP update)

## Files Created/Modified

- `.planning/phases/01-ba-analysis-decision-lock/01-BA-DECISIONS.md` — §4 (90-row inventory), §5 (7 questions + EV-M-only list), §6 (summary + Phase 2 mapping), §8 (K1/K2 sweep appendix) all completed; document status remains DRAFT (business-owner confirmation is Plan 03's job)

## Decisions Made

See `key-decisions` in frontmatter. Summary:
- §4.C rewritten against the real BOM function (`_parse_sheet`) after confirming the plan's cited function is THDM-only.
- D2/D3 (Excel/Bien_doi branches of `_resolve_row_mapping`) verdict BỎ, confirmed dead-for-BOM via exhaustive call-site grep, not inference.
- K3's SỬA target corrected: the transaction-wrapping SQL-06 seemed to be missing is actually already present in `_run_insert_bg`; the real gap is ARCH-03's staging/validation layer.
- BA-02 intentionally left out of `requirements-completed` — it requires a business-owner confirmation this plan cannot self-certify.

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Rewrote §4.C (C1-C10) against the real BOM parse function**
- **Found during:** Task 1, read_first step
- **Issue:** Plan's `read_first`/action text specified group C (row-parse rules) against `bom_parser.py:297-407` (`_parse_section_excel_rows`). Cross-checking against `01-01-SUMMARY.md`'s own corrected finding (EV-S13) and `EV-S08`'s explicit note ("Rule chữ-hoa-toàn-bộ...CHỈ tồn tại trong `_parse_section_excel_rows`... BOM dùng `_parse_sheet`") confirmed this function is THDM-only (grep: called only once, at `bom_parser.py:1143`, from a `_thdm_` function).
- **Fix:** Read `_parse_sheet` (`bom_parser.py:461-588`) directly and rewrote C1-C10 against its real logic, with an explicit baseline note at the top of §4.C explaining the correction and flagging per-row where the actual mechanism differs from the plan's original description (C4, C6, C8 in particular have no analog in the real BOM function).
- **Files modified:** `01-BA-DECISIONS.md` (§4.C)
- **Verification:** Task 1's automated verify (C1-C10 presence + citation checks) passes; every C-row cites `.py:line` pointing at real, confirmed-executing BOM code.
- **Committed in:** `5ba67ea`

**2. [Rule 1 - Bug] Corrected D2/D3 verdicts after confirming the code path is dead for BOM**
- **Found during:** Task 1, §4.D action step
- **Issue:** The plan's action text for D2/D3 assumed `_resolve_row_mapping`'s `Excel`/`Bien_doi` branches apply to BOM (matching `bom_parser.py`'s general docstring). Grepping all call sites of `_resolve_row_mapping` in the repo (exactly 2: `main_window.py:6588` for THDM with `get_excel_val` supplied, `main_window.py:9170` for BOM with `get_excel_val=None` and a pre-filtered `_non_excel` record list) showed these branches structurally cannot execute for BOM.
- **Fix:** Marked D2/D3 BỎ (not GIỮ) with the call-site enumeration as evidence, and cross-referenced the real BOM-equivalent mechanism (`_resolve_detail_row` Pass 1b, rows E5/E6/E7) so Phase 2 doesn't port dead code.
- **Files modified:** `01-BA-DECISIONS.md` (§4.D)
- **Verification:** Task 1's automated verify passes; §5's EV-M-only list explicitly documents these as code-reading-based (not sample-data-based) BỎ verdicts per the evidence rule.
- **Committed in:** `5ba67ea`

**3. [Rule 1 - Bug] Corrected SQL-06's own evidentiary interpretation (K3)**
- **Found during:** Task 3, §4.K action step, read_first for `_run_insert_bg`
- **Issue:** §3's `SQL-06` note (written in Plan 01) speculated "không có transaction bao trùm cả 2 bước" (no transaction covers both header and detail insert) based on 8 historical orphaned `B20BOM` headers with zero details. Reading `_run_insert_bg` (`main_window.py:10839-10874`) directly showed the current code sets `conn.autocommit = False`, executes the header insert and `_generate_bom_details` on the same connection, and calls `conn.commit()` once at the end with `conn.rollback()` on exception — i.e. a single transaction already wraps both steps.
- **Fix:** K3's Bằng chứng cell explicitly documents this correction ("đính chính bằng chứng") instead of silently repeating the stale interpretation, and redirects K3's SỬA target to the actual remaining gap (ARCH-03 staging/validation, treating the existing transaction as the "Transfer" step of that pattern).
- **Files modified:** `01-BA-DECISIONS.md` (§4.K)
- **Verification:** Task 3's automated integrity-gate verify passes; K3 still cites `SQL-06` and both code locations for the corrected claim.
- **Committed in:** `de5e2fc`

---

**Total deviations:** 3 auto-fixed (all Rule 1 — correctness corrections to citations/evidence interpretation caught before they reached the committed document, not scope creep; consistent with the plan's own explicit "re-read the code and correct the line references if they drift" instruction)
**Impact on plan:** All three corrections were necessary for D-02/D-03's "verdict comes from evidence, not code-reading alone, and every citation must resolve to real, executing code" requirement — an uncorrected citation would have propagated a wrong premise into Phase 2's `core/` port.

## Issues Encountered

None. All three tasks' automated `<verify>` scripts pass (including the full 90-row integrity gate in Task 3); `git status --porcelain` shows no modifications under `Tools/` or `v3/Tools/`; `tests/tmp/ba_bom_flow_sweep.py`/`.json` (and the appendix-generation helper script) were deleted after use, confirmed by the integrity gate's own cleanup assertion.

## User Setup Required

None — no external service configuration required. This plan did not need live SQL (per its own `<mcp_tools>` guidance, it worked entirely from Plan 01's evidence base plus code reading); no new SQL-09+ evidence was needed.

## Known Stubs / Intentional Placeholders

The following are **intentional, on-plan placeholders**, not defects:
- §7 (business-owner sign-off): `_(Điền ở Plan 03...)_` — explicitly scoped to Plan 03 per this plan's own "Artifacts this phase produces" section and the document's own §7 wording ("Im lặng, auto-advance hay yolo mode không tính là xác nhận"). Document status remains DRAFT.
- `Q-01` through `Q-07` all carry an empty `**Trả lời:** (chờ business owner)` slot (or "chờ Phase 2" for the one technical question, `Q-04`) — these are inputs to Plan 03, not gaps in this plan's work.

## Next Phase Readiness

- The full 90-row inventory is complete, evidence-backed, and passes the automated integrity gate — Plan 03 can take the document straight to business-owner review without re-deriving anything.
- 6 questions need a business-owner answer (`Q-01`, `Q-02`, `Q-03`, `Q-05`, `Q-06`, `Q-07`); 1 needs a Phase-2 technical follow-up before its verdict can be finalized (`Q-04`, `BOM5.Code` downstream usage).
- 4 SỬA/BỎ rows (D2, D3, D7, E5) rest on `EV-M01` + code-reading evidence only (no direct sample/SQL proof) — flagged explicitly in §5 per the evidence rule; Phase 2 planners should treat these with the same scrutiny as any EV-M-only finding.
- No blockers for Plan 03.

---
*Phase: 01-ba-analysis-decision-lock*
*Completed: 2026-09-25*

## Self-Check: PASSED

- FOUND: `.planning/phases/01-ba-analysis-decision-lock/01-BA-DECISIONS.md`
- FOUND: `.planning/phases/01-ba-analysis-decision-lock/01-02-SUMMARY.md`
- FOUND commit `5ba67ea` (Task 1)
- FOUND commit `aafa893` (Task 2)
- FOUND commit `de5e2fc` (Task 3)
