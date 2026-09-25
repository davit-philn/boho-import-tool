---
phase: 01-ba-analysis-decision-lock
plan: 01
subsystem: documentation
tags: [bom, mssql-server, bravo-erp, evidence-audit, decision-document, bom-parser]

# Dependency graph
requires: []
provides:
  - "01-BA-DECISIONS.md (status DRAFT): §0 scope/conventions, §1 full commit history 73081dc..8a07639 (33 commits classified + 9 direction-change chains), §2 15 rule-level metrics (EV-S01..EV-S14) + mapping census (EV-M01) over all 94 real BOM sample files, §3 live SQL evidence SQL-00..SQL-08 via mssql-boho, §4 skeleton with tracer row G1 fully evidenced (verdict SỬA, data-driven)"
  - "Proven evidence-ID citation format (EV-Snn / EV-Mnn / SQL-nn) that Plan 02 will cite directly when writing ~90 rule-level verdict rows"
affects: [01-02-rule-level-inventory, 01-03-business-owner-signoff]

# Actuals (#2632)
actuals:
  tokens: 15900
  tasks: 3
  commits: 3

tech-stack:
  added: []
  patterns:
    - "Evidence citation IDs (EV-Snn=94-sample metric, EV-Mnn=mapping census, SQL-nn=live query) as level-3 headings, grep-findable, cited from §4 verdict rows"
    - "One-off evidence-gathering script pattern: import the real parser functions directly (parse_bom_file, _detect_sheet_type, _parse_sheet, etc.) rather than re-deriving predicates from scratch; replicate only the logic that is inline/tied to `self` in main_window.py, with source-line citations"

key-files:
  created:
    - .planning/phases/01-ba-analysis-decision-lock/01-BA-DECISIONS.md
  modified: []

key-decisions:
  - "mssql-boho MCP tool is not exposed to the executor subagent (same class of issue as upstream anthropics/claude-code#13898) even though the server connects at orchestrator level. Resolved by having the coordinator run all SQL-00..SQL-08 queries in the orchestrator session and hand the executor verbatim query text + results to transcribe into §3 — documented explicitly in §0.3 so the evidence chain stays auditable."
  - "Corrected a planning-time assumption: the plan's read_first cited bom_parser.py:378-383 (the Fill-Forward keep-row-value branch) as BOM's Fill-Forward logic, but that function (_parse_section_excel_rows) is THDM-only. BOM uses _parse_sheet (no Fill-Forward merge at parse time at all) plus an unconditional insert-time overwrite in main_window.py:9823-9826. EV-S13 documents the corrected mechanism and its consequence (ItemType's per-row STT resolve is dead code in practice)."
  - "G1 (_mkt_cache dict-collision) verdict: SỬA, driven by live SQL-01 data (4 real collision groups in vB20Item_MKT, nondeterministic winner across app runs) — not by code-reading inference alone, per D-03. The concrete fix (priority rule) is deferred to Q-01 for the business owner, not chosen by the executor."

requirements-completed: [BA-01, BA-03]

coverage:
  - id: D1
    description: "01-BA-DECISIONS.md skeleton (§0-§8 headings, 5-column table convention) with tracer row G1 evidenced end-to-end (code + git + live SQL)"
    requirement: BA-03
    verification:
      - kind: other
        ref: "Task 1 <verify> automated script embedded in 01-01-PLAN.md (heading/status/G1-row structure assertions)"
        status: pass
    human_judgment: false
  - id: D2
    description: "§1 full changelog (D-04, 33 commits classified + 9 direction-change chains) and §2 evidence base (D-02, EV-S01..EV-S14 + EV-M01) over all 94 real sample files"
    requirement: BA-01
    verification:
      - kind: other
        ref: "Task 2 <verify> automated script embedded in 01-01-PLAN.md (commit-hash presence, EV-S00..EV-S14/EV-M01 heading + x/94 + sample-filename assertions)"
        status: pass
    human_judgment: false
  - id: D3
    description: "§3 live SQL evidence SQL-00..SQL-08 (D-03), all read-only SELECT/WITH/catalog queries with verbatim query + result, tests/tmp/ evidence scripts deleted per cleanup step"
    requirement: BA-01
    verification:
      - kind: other
        ref: "Task 3 <verify> automated script embedded in 01-01-PLAN.md (SQL-00..SQL-08 heading presence, no data-modifying keywords, result-line count, temp-file-deleted assertion)"
        status: pass
    human_judgment: false

duration: 1h 49min
completed: 2026-09-25
status: complete
---

# Phase 1 Plan 01: BA Evidence Base & Tracer Summary

**Full BOM evidence base (94-sample metrics + 33-commit changelog + 9 live SQL queries) with tracer row G1 proving the `_mkt_cache` dict-collision hypothesis confirmed by real production data.**

## Performance

- **Duration:** 1h 49min (2026-09-25T04:44:18Z → 2026-09-25T06:33:04Z, spanned a rate-limit pause and resume mid-Task-2)
- **Started:** 2026-09-25T04:44:18Z
- **Completed:** 2026-09-25T06:33:04Z
- **Tasks:** 3/3
- **Files modified:** 1 (`01-BA-DECISIONS.md`, created)

## Accomplishments

- **Tracer row G1 proven end-to-end**: code (`main_window.py:9529-9538`, `9835-9839`) + git (`e92c060`, the only commit that ever touched `_mkt_cache`, never revisited since v2.2.4) + live SQL (`SQL-01`: 4 real `ItemTypeSX_Parent` collision groups in `vB20Item_MKT`, and `SQL-01(d)`/`SQL-07` showing the "winning" Id is nondeterministic across app runs and accounts for 84% of all real MKT-fallback traffic ever). Verdict: **SỬA**, Mức rủi ro: **Cao** — driven entirely by the data, per D-03.
- **Full changelog (D-04)**: 33 commits from `73081dc` (v2.2.1) to `8a07639` (v2.2.25) classified FIX/FEAT/ĐỔI HƯỚNG/REVERT and linked to row IDs, plus 9 direction-change chains (8 specified by the plan + 1 found during execution: the v2.2.20/v2.2.21 DETAIL label-mapping fix was applied to BOM2 but never to BOM5, directly explaining a real EV-S04 finding and corroborating `docs/audit/AUDIT-2026-09-21.md` finding #3).
- **15 rule-level metrics + mapping census (D-02)**, all measured on the real 94 sample files (not simulated): EV-S01 through EV-S14 plus EV-M01, each with an x/94 count and a concrete sample filename. Notable findings: BOM's `_parse_sheet` has no all-caps-group-title drop rule at all (that rule only exists in the THDM-only pipeline — 0/94 is structural, not a measurement gap); BTP Rule 1b never occurs alone in the sample set (always subsumed by Rule 1a); `ItemType`'s per-row resolve from Excel STT is dead code in practice because insert-time Fill-Forward unconditionally overwrites it.
- **8 live SQL queries (SQL-02..SQL-08, D-03)** covering structural risks the Excel samples cannot prove: BTP code-reuse collision scope (deterministic, unlike G1), the `usp_B20BOM_Create_ItemCode` synonym chain across two databases, post-fix residual error rates for `EmployeeId` and `ParentDetailRowId_SO` (both nonzero — flagged as open questions, not assumed resolved), orphaned-import evidence (all historical, no active problem), and the quantitative basis (84% of MKT-fallback traffic) for G1's risk rating.

## Task Commits

Each task was committed atomically:

1. **Task 1 (tracer): skeleton + G1 end-to-end** - `d56112d` (feat)
2. **Task 2: full changelog + 94-sample evidence base** - `6c757c7` (docs)
3. **Task 3: SQL-02..SQL-08 + cleanup** - `a8bf8eb` (docs)

**Plan metadata:** commit pending (this SUMMARY + STATE/ROADMAP update)

## Files Created/Modified

- `.planning/phases/01-ba-analysis-decision-lock/01-BA-DECISIONS.md` - the single Phase 1 deliverable; §0-§3 complete (status DRAFT), §4 skeleton with tracer row G1 filled, §5-§8 headings present for Plan 02/03

## Decisions Made

See `key-decisions` in frontmatter. Summary:
- mssql-boho MCP unavailable to this executor subagent → coordinator ran all §3 queries in the orchestrator session; verbatim transcription documented in §0.3 for auditability.
- Corrected a planning-time assumption about which function implements BOM's Fill-Forward (it's `main_window.py`'s insert-time step, not `bom_parser.py`'s THDM-only `_parse_section_excel_rows`) — captured as EV-S13.
- G1 verdict follows the data (SỬA), with the specific priority-rule fix deferred to the business owner as `Q-01`, per D-03/D-05.

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Corrected Fill-Forward predicate source before writing EV-S13**
- **Found during:** Task 2 (§2, EV-S13 action step)
- **Issue:** Plan's read_first cited `bom_parser.py:378-383` as the Fill-Forward "keep row value unless override" logic for BOM. That function (`_parse_section_excel_rows`) is exclusively the THDM pipeline; BOM uses `_parse_sheet`, which has no Fill-Forward merge at the parse layer at all.
- **Fix:** Traced the actual BOM Fill-Forward mechanism (insert-time, `main_window.py:9800-9806` capture + `9823-9826` unconditional overwrite) and wrote EV-S13 against the real code path, explicitly noting the corrected understanding so Plan 02 doesn't inherit the wrong citation.
- **Files modified:** `01-BA-DECISIONS.md` (§2, EV-S13)
- **Verification:** Cross-checked against `main_window.py` line ranges read directly; Task 2 automated verify passes.
- **Committed in:** `6c757c7`

**2. [Rule 1 - Bug] Fixed evidence-gathering script bugs found via cross-validation (not app bugs)**
- **Found during:** Task 2, while building `tests/tmp/ba_sample_evidence.py`
- **Issue:** Two bugs in the temporary evidence script itself (not the production app): (a) `Tên vật tư`/`Tên chi tiết` column lookup used exact-match only, missing the real suffix-merged header pattern (`SLg_Tên_Vật_Tư`) that the app itself resolves via a 2-pass exact/suffix match — this silently zeroed out EV-S09/EV-S10 on most files; (b) the main evidence pass called `parse_bom_file(path)` without `meta_keys`/`cell_specs` arguments, so `global_meta` came back empty for EV-S14 (Y3/Số lượng presence).
- **Fix:** Added a `find_col_exact_or_suffix()` helper matching the app's real 2-pass logic; added a dedicated supplementary script that calls `parse_bom_file()` with the correct arguments (matching how the real app's `build_meta_keys_from_mapping`/`build_cell_specs_from_mapping` are meant to be used) to recompute EV-S14 accurately.
- **Files modified:** `tests/tmp/ba_sample_evidence.py` (temporary, deleted in Task 3 per plan)
- **Verification:** Re-ran and cross-checked counts (e.g. suffix-match jumped from 0 to 71/94 files; EV-S14 confirmed 94/94 via the corrected supplementary script) before writing final numbers into the document.
- **Committed in:** `6c757c7` (only the corrected final numbers reached the committed document; the script itself was deleted in `a8bf8eb`)

---

**Total deviations:** 2 auto-fixed (both Rule 1 - correctness bugs caught before they reached the committed evidence, not scope creep)
**Impact on plan:** Both fixes were necessary for D-02/D-03's "no inference without real cross-referencing" requirement — an uncorrected citation or an undercounted metric would have propagated a wrong premise into Plan 02's ~90-row inventory.

## Issues Encountered

- **mssql-boho MCP tool not exposed to this executor subagent** (server connected, but `mcp__mssql-boho__*` tools returned "connected but does not offer this tool here" for every name tried, and `ToolSearch` was disabled for the session). This is the precondition both Task 1 and Task 3 declared. Per the execution protocol, an unmet precondition is never auto-approved — the plan halted with a `checkpoint:human-verify` after Task 1's non-SQL prerequisites but before any SQL evidence was fabricated. The coordinator resolved this by running the queries in the orchestrator session (where the tool works) and relaying verbatim query text + results for transcription, which is the exact "hand me the verbatim results" resolution path proposed in the checkpoint. Documented in §0.3 of the decision document for full auditability of the evidence chain.
- No other blockers. All three tasks' automated `<verify>` scripts pass; `git status --porcelain` shows no modifications under `Tools/` or `v3/Tools/`.

## User Setup Required

None - no external service configuration required (the `mssql-boho` connectivity gap was a subagent tool-exposure limitation, not a missing user credential; see Issues Encountered).

## Known Stubs / Intentional Placeholders

The following are **intentional, on-plan placeholders**, not defects — the plan's own output spec assigns them to later plans and they are not appended to the cross-phase WINDOWS.md ledger:
- §4.A-§4.F, §4.H-§4.K (all BTP/header/resolve rows except tracer row G1), §6, §8: `_(Điền ở Plan 02.)_` — explicitly scoped to Plan 02 per `01-01-PLAN.md`'s "Artifacts this phase produces" section.
- §7 (business-owner sign-off): `_(Điền ở Plan 03...)_` — explicitly scoped to Plan 03.
- EV-S11's non-str-value count is a documented **lower bound** (capped at 5,000 during collection to keep the evidence JSON manageable; true count is higher) — noted inline in the document itself, not hidden.

## Next Phase Readiness

- Evidence base (§1-§3) is complete and citable — Plan 02 can write ~90 rule-level verdict rows (§4.A-§4.F, §4.H-§4.K) directly against `EV-Snn`/`EV-Mnn`/`SQL-nn` IDs without re-deriving evidence.
- Four open questions recorded for the business owner / Phase 2 in §5 (`Q-01` MKT priority rule, `Q-02` EmployeeId=1 post-fix legitimacy, `Q-03` ParentDetailRowId_SO residual errors, `Q-04` BOM5.Code non-str downstream usage) — Plan 02/03 should either answer these from further code reading (Q-04) or explicitly route them to the business owner (Q-01/Q-02/Q-03) rather than silently assuming an answer.
- No blockers for Plan 02.

---
*Phase: 01-ba-analysis-decision-lock*
*Completed: 2026-09-25*

## Self-Check: PASSED

- FOUND: `.planning/phases/01-ba-analysis-decision-lock/01-BA-DECISIONS.md`
- FOUND: `.planning/phases/01-ba-analysis-decision-lock/01-01-SUMMARY.md`
- FOUND commit `d56112d` (Task 1)
- FOUND commit `6c757c7` (Task 2)
- FOUND commit `a8bf8eb` (Task 3)
