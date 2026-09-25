# Phase 1: BA Analysis & Decision Lock - Pattern Map

**Mapped:** 2026-09-25
**Files analyzed (new/modified):** 0 source files — this phase produces one decision document, no code changes
**Analogs found:** N/A (document-structure analogs only)

## Scope Note

Phase 1 does not create or modify any source code. Per `01-CONTEXT.md` (`<domain>`), the only deliverable is a decision document (5-column table: `Logic/rule con | Verdict | Bằng chứng | Cách sửa | Mức rủi ro`) inventorying existing BOM logic. No `core/` code is written until Phase 2. Consequently there is no "File Classification" / "Pattern Assignments" table of new files — instead this document maps (a) the closest existing document-format analog for the deliverable itself, and (b) the shape/location of the source code the analysis must read.

## Document-Format Analog

**Deliverable:** `01-DECISIONS.md` (or similarly named) — 5-column BOM logic inventory table

**Closest analog:** `.planning/codebase/CONCERNS.md` (audit doc, 2026-08-03)

**Why it's the best analog:** Same genre of artifact — an evidence-based engineering inventory produced from reading the same codebase (`bom_parser.py`, `main_window.py`), organized as repeated per-item blocks each with Issue/Files/Impact/Fix-approach. This maps directly onto Phase 1's required 5 columns:

| CONCERNS.md field | Phase 1 column |
|---|---|
| Issue (heading + description) | Logic/rule con |
| (implicit — codebase read) | Verdict (GIỮ/SỬA/BỎ) |
| Files + line refs, quoted code | Bằng chứng |
| Fix approach (with before/after code) | Cách sửa (nếu SỬA) |
| Impact | Mức rủi ro |

**Concrete excerpt to follow (CONCERNS.md lines 13-22):**
```
**Large Complex Functions in bom_parser.py:**
- Issue: Multiple functions exceed 100+ lines with deep nesting...
  - `_resolve_row_mapping()` — 115 lines
- Files: `Tools/services/bom_parser.py`
- Impact: Hard to debug, test individual paths...
- Fix approach: Extract sub-functions for each logical step...
```
This confirms two things the Phase 1 doc should reuse: (1) always cite exact function names + line/length, (2) pair every claim with a `Files:` pointer — directly satisfying D-02/D-03's evidence requirement.

**Also relevant:** `.planning/codebase/ARCHITECTURE.md` and `STRUCTURE.md` — read only for extra module-boundary context if needed (per CONTEXT.md canonical_refs); they are not structural analogs for the decision table itself (they're narrative architecture docs, not verdict tables).

## Source Files the Analysis Must Read (targets of inventory, not files to modify)

| File | Role in inventory | Key regions (from CONTEXT.md canonical_refs) |
|------|--------------------|-----------------------------------------------|
| `v3/Tools/services/bom_parser.py` | Excel parser — section/header detection, Fill-Forward, `_parse_section_excel_rows`, `_resolve_row_mapping` | whole file (already flagged in CONCERNS.md as containing `_resolve_row_mapping()` @ 115 lines, `_parse_section_excel_rows()` @ 113 lines) |
| `v3/Tools/views/main_window.py` | UI-coupled resolve logic | lines 9127-9250+ (`_resolve_detail_row`); lines 9490-9890 (`_build_bom_detail_caches`/BTP detection/`_mkt_cache`/SP_HOOK insert loop); lines 4780-4820, 8940-8945 (`_current_creator_employee_id`) |
| `v3/Tools/services/mapping_loader.py` | Config structure BOM logic depends on | `_CONFIG`, `SP_HOOK`, `SP_CONFIG` |
| `v3/docs/audit/bom_sample_scan.json` | Evidence source (94 sample files scan results) | whole file — used for D-02 evidence citations |

No Grep/Glob-based "analog controller/service" search was performed beyond the above, because this phase writes zero source files — role/data-flow classification (controller, service, CRUD, etc.) does not apply.

## Shared Patterns to Carry Forward

### Evidence citation format (from CONCERNS.md)
Every SỬA/BỎ verdict row should cite: exact function name, line number or line range, and a `Files:` path — mirroring CONCERNS.md's convention. Where D-02 evidence is a sample file, cite the specific filename from `v3/sample_data/bom_import/` and the relevant field in `bom_sample_scan.json`. Where D-03 evidence is a live query, cite the exact SQL run via `mssql-boho` MCP and its result shape.

### Rule-level granularity (from D-01, reinforced by CONCERNS.md's per-function breakdown style)
CONCERNS.md already itemizes at function granularity (e.g., listing 5 named functions individually rather than saying "parser has long functions"). Phase 1 must go one level deeper per D-01 — one row per rule/branch (e.g., BTP detection Rule 1 vs Rule 2; each `NguonDL` source in `_resolve_row_mapping`/`_resolve_detail_row` evaluated separately) rather than one row per function.

## No Analog Found

None — the phase's only output (a markdown decision table) has a direct structural analog in `CONCERNS.md`, and all source files to be analyzed are explicitly named in `01-CONTEXT.md`.

## Metadata

**Analog search scope:** `.planning/codebase/*.md` (7 files); read `CONCERNS.md` in full for structural pattern.
**Files scanned:** 1 read in full (CONCERNS.md), 6 globbed and referenced by name (ARCHITECTURE.md, STRUCTURE.md, STACK.md, CONVENTIONS.md, TESTING.md, INTEGRATIONS.md — not read, not needed for this phase's document-format question).
**Pattern extraction date:** 2026-09-25
