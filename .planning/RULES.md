# VIBE CODING SYSTEM RULES

## 1. SCOPE & SAFETY
- Surgical Precision: Only modify code inside the target scope. Do NOT refactor working unrelated code.
- DB Safety: NO DROP/TRUNCATE. Preview UPDATE/DELETE SQL before executing via MCP.
- Circuit Breaker: Stop after 3 failed test attempts. Summarize the blocker and ask the user.

## 2. TOKEN & WORKFLOW OPTIMIZATION
- Heavy Sweeps: Never use repetitive read_file/ripgrep > 5 times. Write a temporary Python script to dump report to JSON/TXT, then read the report.
- Auto Checkpoint: Run `git commit` automatically whenever `pytest` achieves 100% pass rate.
- Workspace Cleanliness: Store temp files in `tests/tmp/` and auto-delete them before declaring DONE.

## 3. RESPONSE STYLE
- Executive Summary: Concise output only. Status = What was done -> Test/SQL Verification result -> Next step. No fluff.