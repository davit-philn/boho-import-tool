# UI Review — BOHO IMPORT BOM/THDM (CustomTkinter Desktop App)

**Audited:** 2026-08-08
**Baseline:** Abstract 6-pillar standards (no UI-SPEC.md present)
**Screenshots:** Not captured (no dev server — desktop Tkinter app, code-only audit)
**App type:** CustomTkinter + Tkinter hybrid, Windows 11, Vietnamese-language internal tool

---

## Pillar Scores

| Pillar | Score | Key Finding |
|--------|-------|-------------|
| 1. Copywriting | 3/4 | Most Vietnamese copy is accurate, but mapping info label (line 1480) falls back to unaccented/mixed-language text |
| 2. Visuals | 3/4 | Strong hierarchy with numbered workflow steps; settings button is icon-only with no accessible tooltip or aria equivalent |
| 3. Color | 2/4 | THEMES system is well-designed but bypassed 178+ times via hardcoded hex; empty-state text colors are dark-mode-only and will render illegible in Light mode |
| 4. Typography | 3/4 | Clean 7-level scale from constants; update popup and sidebar label bypass the scale with raw tuples including non-standard 9pt and 36pt sizes |
| 5. Spacing | 3/4 | PAD_XS/SM/MD/LG tokens exist and are used in major areas, but magic numbers 6, 10, 14, 22, 24, 26, 28 appear throughout without token references |
| 6. Experience Design | 4/4 | Excellent: startup retry, loading popups, empty states with numbered workflow, warning panel, toast system, copy feedback, keyboard shortcuts |

**Overall: 18/24**

---

## Top 3 Priority Fixes

1. **Empty-state text uses hardcoded dark-only hex colors** (`fg="#3E3E42"` and `fg="#555555"` at lines 1155, 1159, 1163) — in Light mode these near-black values render nearly invisible against the white `sheet_table_bg`; fix by replacing with `dt["text_muted"]` and `dt["text_main"]` from the theme dict (already computed at that point in the code as `_EMPTY_BG = dt["sheet_table_bg"]`).

2. **`evenrow` tag hardcodes dark-mode panel color** (`background="#2A2D2E"` at line 1193) — this tag is never updated in `apply_theme()`, so switching to Light mode leaves alternating rows with a very dark background against white text; fix by pulling the value from `dt["bg_panel"]` at build time and re-tagging in `apply_theme()`.

3. **Mapping info label loses Vietnamese diacritics** (line 1480-1484: `"Tong: "`, `"Can xac nhan:"`, `"System:"`) — internal staff sees this label daily; fix by replacing with `"Tổng:"`, `"Cần xác nhận:"`, `"Hệ thống:"` to match the rest of the UI's copy standard.

---

## Detailed Findings

### Pillar 1: Copywriting (3/4)

**Passes:**
- Numbered workflow prompts (①②③④⑤) consistently guide users through each tab.
- Custom dialogs use natural Vietnamese: "Có"/"Không", "Đóng", "Thử lại", "Bỏ qua".
- Error messages are specific and actionable: `"⚠️  Kiểm tra VPN hoặc cấu hình DB rồi thử lại"` (line 784), `"Lỗi khi tải. Kiểm tra kết nối mạng."` (line 253).
- Status labels use semantic color differentiation: green tick for success, red X for errors, hourglass for loading.
- Toast messages and status labels are contextual, not generic.

**Failures:**
- WARNING: `_map_info_lbl` at line 1480-1484 reverts to unaccented/mixed-language Vietnamese: `"Tong: "`, `"Can xac nhan:"`, `"System:"`. This is the Mapping Config panel's footer summary. It contradicts the diacritics standard of every other label in the app.
- MINOR: `"Đang tải... —"` placeholder in combos is clear, but the em dash `—` in the placeholder feels odd; `"— Đang tải —"` would be more consistent with the `"— Chọn nhân viên —"` pattern.
- MINOR: `"Bỏ qua →"` on startup failure (line 794) uses a raw Unicode arrow `→` instead of standard button text. Not wrong, but slightly inconsistent with emoji-icon pattern used everywhere else.
- MINOR: The right-click context menu label `"Copy ô đang click        (click trái)"` (line 1356) has excessive whitespace used for column alignment — this alignment is tab-stop dependent and will misalign on different font renderers.

---

### Pillar 2: Visuals (3/4)

**Passes:**
- Strong focal-point design on each empty state: large 48pt emoji icon (📂, 📋) as the visual anchor, then a primary instruction, then a secondary hint.
- Tab icons (📥, 📊, 🌳) give instant categorical differentiation.
- The update popup has a clear visual hierarchy: bell icon (36pt) → title (14pt bold) → notes (10pt) → progress bar → CTA button.
- Startup overlay uses a circular branded "B" badge as the product identity mark.
- Warning panel uses badge + collapsible list — a Slack-style pattern that suits the data-dense context.
- Separator lines between toolbar sections (thin 1px frames) cleanly divide action groups.

**Failures:**
- WARNING: `self._btn_settings` (line 889) displays only `"🎨"` (an emoji) with no text and no `Tooltip()` attached. Every other icon-only element in the app (`lbl_status` at line 874, `btn_undo_import` at line 1036) has a `Tooltip()`. The settings button has no discoverable affordance beyond visual trial.
- MINOR: The sidebar sheet list (`BẢNG DỮ LIỆU` label, line 1052) uses `("Segoe UI", 9, "bold")` — the smallest possible font at 9pt. On high-DPI displays this label may be illegible. The minimum body size in the defined scale is 10pt (FONT_SMALL).

---

### Pillar 3: Color (2/4)

**Passes:**
- `THEMES` dict in `utils.py` is comprehensive: 40+ semantic tokens per theme covering backgrounds, text, borders, accents, semantic status (log_ok, log_warn, log_err), badge colors, scrollbars, and popup backgrounds.
- Light mode palette uses proper inversion: `bg_main="#F8FAFC"` / `text_main="#0F172A"` — not simply a brightness invert.
- Semantic color is used correctly: `dt["log_ok"]` for success, `dt["log_err"]` for errors, `dt["badge_warn_bg"]` for warnings.
- `apply_theme()` builds a `_color_map` to recolor legacy `tk.Frame`/`tk.Label` widgets that cannot observe CTk appearance changes — good defensive engineering.

**Failures:**
- BLOCKER: `fg="#3E3E42"` and `fg="#555555"` at lines 1155, 1159, 1163 (empty-state labels) are dark-only values. In Light mode, `sheet_table_bg="#FFFFFF"` — near-black text on white background has high contrast, so the labels are technically readable, BUT the icon emoji label at line 1155 uses `fg="#3E3E42"` as foreground color for a tk.Label that displays an emoji. More critically, these labels are created at build time with the current-theme value and are NOT added to `_recolor_tk_widgets` remapping — so `apply_theme()` will not update them.
- BLOCKER: `self.tree.tag_configure("evenrow", background="#2A2D2E")` at line 1193. This is a raw dark color never registered in `THEMES` and never updated in `apply_theme()`. In Light mode, alternating data rows will have a very dark gray background with light text: near-unreadable contrast.
- WARNING: `_warn_tog = "#888888"` at line 521 is hardcoded during `apply_theme()`. It does not differentiate between Dark and Light. In Light mode `#888888` on `danger_hdr_bg="#FFF0F0"` is acceptable, but the approach bypasses the token system.
- WARNING: `tip.configure(bg="#094771")` at line 1260 (copy tooltip) and `activebackground="#094771"` at line 1357 (right-click menu) are hardcoded VS Code accent blues — not in `THEMES`. In Light mode the copy tooltip may not match the overall palette.
- WARNING: `lbl_status` in the update popup (line 222) hardcodes `fg="#4FC3F7"` (the Dark mode accent color). This is inside a `tk.Toplevel` built with `bg = "#1E1E1E" if is_dark else "#F5F5F5"` but the `fg` is never adjusted for Light mode.
- INFO: `background="#2A2D2E"` evenrow is closely related to `bg_panel="#252526"` in Dark but uses a slightly different value — minor inconsistency even within Dark mode.
- INFO: Total raw hex values in main_window.py: 178. The majority are in UI definition code, not the theme remap dict. The THEMES system covers mapping but cannot auto-capture new widget usages that don't go through `_recolor_tk_widgets`.

---

### Pillar 4: Typography (3/4)

**Passes:**
- Seven-level type scale is well-defined as named constants in `utils.py`:
  - FONT_SMALL (10), FONT_BODY (11), FONT_MD (12), FONT_LABEL (13), FONT_TITLE (18), FONT_ICON (42), FONT_HERO (48)
- Both normal and bold variants exist for each size (FONT_BODY/FONT_BODY_B, FONT_MD/FONT_MD_B, etc.).
- Italic variant (`FONT_SMALL_I`) is used sparingly for "inactive" catalog items — correct semantic use.
- Constants are imported and used throughout, making future font changes a single-point edit.

**Failures:**
- WARNING: The force-update popup (`_on_update_found`, lines 216-219) uses three raw font tuples:
  - `("Segoe UI", 36)` — bell icon (no constant; closest would be between FONT_ICON=42 and FONT_TITLE=18)
  - `("Segoe UI", 14, "bold")` — "Có phiên bản mới" (no constant; gap between FONT_LABEL=13 and FONT_TITLE=18)
  - `("Segoe UI", 10)` = FONT_SMALL, but written as a raw tuple rather than the constant
  - These hardcodings mean the update popup does not benefit from any future type-scale changes.
- WARNING: Line 1054 sidebar section header uses `font=("Segoe UI", 9, "bold")` — 9pt is below FONT_SMALL (10pt) and outside the defined type scale. On a 1360px-wide display this label is `"BẢNG DỮ LIỆU"` at 9pt — borderline legible.
- INFO: Seven font sizes is at the upper end of recommended hierarchy depth (typically 4–5 for a single-screen app). The FONT_ICON (42) and FONT_HERO (48) are only used in empty states and the startup overlay logo, so in practice most screens use 5 levels — acceptable.

---

### Pillar 5: Spacing (3/4)

**Passes:**
- Four-step spacing scale (`PAD_XS=4`, `PAD_SM=8`, `PAD_MD=12`, `PAD_LG=20`) is defined in `utils.py` and imported.
- Constants are used for major structural padding: settings dialog (`padx=PAD_LG`), top bar (`pady=PAD_MD`), button groups (`padx=PAD_SM`), toolbar separators, dialog button rows.
- Toolbar button height is consistently `height=32` across the action bar.
- Dialog corner radii are consistent (`corner_radius=8` for primary CTAs, `corner_radius=6` for secondary and small buttons).

**Failures:**
- WARNING: Magic numbers appear alongside tokens without clear rationale:
  - `pady=10` (lines 965, 998, 1015, 1029) — 10 is between PAD_SM(8) and PAD_MD(12). No token.
  - `padx=6` (lines 1058, 1112, 1414) — 6 is between PAD_XS(4) and PAD_SM(8). No token.
  - `padx=28` (lines 705, 708) — close to PAD_LG(20) but not equal.
  - `ipadx=24, ipady=4` (line 702) — startup DB card internal padding, arbitrary.
  - `pady=(0, 26)` (line 702) — between PAD_LG(20) and 32. No token.
  - `padx=10` (line 1448) — close to PAD_MD(12). No token.
- INFO: The inconsistency between `pady=10` (toolbar) and `pady=PAD_SM` (8) or `pady=PAD_MD` (12) produces a slight vertical misalignment in button rows where both appear.

---

### Pillar 6: Experience Design (4/4)

**Passes:**
- Startup flow: DB connection check with overlay, auto-retry 3 times (2s apart), then user-facing retry + skip buttons with specific error text (lines 723-799). Excellent — prevents user from landing in a broken state silently.
- Loading states: `_make_loading_popup()` with indeterminate progress bar blocks the window and prevents double-submit. Loading popup cannot be closed by the user (WM_DELETE_WINDOW lambda: None, line 9544). Updates to `_loading_lbl` from background threads are dispatched via `self.after(0, ...)` correctly.
- Force update: `_on_update_found()` creates a blocking, unclosable dialog with real progress bar showing download percentage, transitions to "applying" state, then triggers restart — complete and correct flow.
- Empty states: Three distinct empty-state contexts (Import tab, THDM tab, catalog tab) each display numbered workflow steps specific to their context. Users always know what action to take next.
- Warning panel: Collapsible error list with row-count badge, toggle arrow, click-to-navigate from error list to corresponding data row. Better than the typical "N errors found" modal.
- Disabled states: Sequential workflow enforced — "Kiểm tra" disabled until file loaded, "Import" disabled until validation passes, "Hoàn tác" disabled until import succeeds.
- Copy feedback: Click-to-copy on any cell with 500ms highlight flash and `"✓ Đã copy: ..."` tooltip (auto-dismissing after 1.4s). Right-click context menu covers all clipboard scenarios.
- Toast system (`_toast()`): Non-blocking, bottom-center, auto-dismissing, 4 semantic kinds with color coding.
- Keyboard shortcuts: Return/Escape on all dialogs, Ctrl+C, Ctrl+A on all tables.
- Tooltip system: Applied to status labels and destructive-action buttons explaining prerequisites.

**Minor gaps (do not affect score):**
- The settings popup anchors to the `_btn_settings` button position but the button has no `Tooltip()` explaining what "🎨" does. A new user would not know this opens display settings.
- The `_ask_msg` dialog binds Return to "Có" (Yes) — this is the more dangerous default for destructive confirmation dialogs. Consider making Return only confirm when the "Có" button has focus.

---

## Files Audited

- `D:\project\insertdatasqlserver\Tools\views\main_window.py` (495KB — audited via offset/limit reads covering lines 1–1600, 4400–5400, 9406–9560)
- `D:\project\insertdatasqlserver\Tools\services\utils.py` (full file — THEMES palette, font constants, spacing tokens)
