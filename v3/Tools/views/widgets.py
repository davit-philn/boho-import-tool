"""
views/widgets.py — Reusable UI widget classes (CustomTkinter + tksheet).
"""
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
try:
    import tksheet
    _HAS_TKSHEET = True
except ImportError:
    _HAS_TKSHEET = False

from services.utils import guess_col_align, THEMES

# ── Compat wrappers: business logic gọi .config(fg=, bg=) như tkinter thường ─
class CLabel(ctk.CTkLabel):
    def config(self, **kw):
        mp = {}
        for k, v in kw.items():
            if k in ("fg", "foreground"):   mp["text_color"] = v
            elif k in ("bg", "background"): mp["fg_color"]   = v
            else:                           mp[k]            = v
        super().configure(**mp)

class CButton(ctk.CTkButton):
    def config(self, **kw):
        mp = {}
        for k, v in kw.items():
            if k in ("fg", "foreground"):   mp["text_color"] = v
            elif k in ("bg", "background"): mp["fg_color"]   = v
            else:                           mp[k]            = v
        super().configure(**mp)


# ── SearchCombo: Entry + popup Listbox có ô tìm kiếm (kiểu BRAVO) ─────────────
class _SearchCombo(ctk.CTkFrame):
    """Searchable dropdown: hiển thị Entry, click/nhập mở popup Listbox có search."""

    _BG_DARK  = "#252526"
    _BG_ENTRY = "#3C3C3C"
    _FG       = "#E8E8E8"
    _SEL_BG   = "#0066CC"
    _SEL_FG   = "#FFFFFF"
    _FONT     = ("Segoe UI", 12)
    _BORDER   = "#3D3D3D"

    # Combo đang mở popup (class-level) — mở combo mới thì đóng combo cũ
    _open_combo = None

    def __init__(self, master, values=None, width=200, height=32,
                 font=None, command=None, placeholder="—", state="normal",
                 popup_width=None, **kw):
        super().__init__(master, fg_color="transparent",
                         width=width, height=height, **kw)
        self.pack_propagate(False)
        self._width      = width
        self._height     = height
        self._popup_width = popup_width  # None = auto (max widget_width, 480)
        self._font       = font or ctk.CTkFont("Segoe UI", 11)
        self._command    = command
        self._placeholder = placeholder
        self._all_values: list[str] = list(values or [])
        self._popup: tk.Toplevel | None = None
        self._listbox: tk.Listbox | None = None
        self._pop_search_var = tk.StringVar()
        self._state  = state          # "normal" | "disabled"
        self._value  = tk.StringVar(value=placeholder)

        # ── Display entry (read-only, click → popup) ──
        self._entry = ctk.CTkEntry(
            self, textvariable=self._value,
            font=self._font, height=height,
            corner_radius=6, width=width - 30,
            state="readonly" if state == "normal" else "disabled",
            fg_color=("#FFFFFF", "#2D2D30"),
            text_color=("#0F172A", "#E8E8E8"),
            border_color=("#E2E8F0", "#444444"))
        self._entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._entry.bind("<Button-1>", self._toggle_popup)

        # ── Arrow button ──
        self._btn = ctk.CTkButton(
            self, text="▾", width=28, height=height,
            corner_radius=6,
            fg_color=("#F1F5F9", "#2D2D30"),
            hover_color=("#E2E8F0", "#3C3C3C"),
            text_color=("#64748B", "#AAAAAA"),
            font=ctk.CTkFont("Segoe UI", 10),
            command=self._toggle_popup,
            state=state)
        self._btn.pack(side=tk.LEFT)

    @classmethod
    def update_theme(cls, t: dict):
        """Cập nhật class-level colors cho popup — áp ngay cho popup mở tiếp theo."""
        cls._BG_DARK  = t.get("popup_bg",     cls._BG_DARK)
        cls._BG_ENTRY = t.get("popup_entry_bg", cls._BG_ENTRY)
        cls._FG       = t.get("text_main",    cls._FG)
        cls._SEL_BG   = t.get("sel_bg",       cls._SEL_BG)
        cls._BORDER   = t.get("popup_border", cls._BORDER)

    # ── Public API (compatible với CTkComboBox) ───────────────────────────────
    def get(self) -> str:
        return self._value.get()

    def set(self, value: str):
        self._value.set(value)

    def set_values(self, values: list[str], reset_selection: bool = True):
        """Cập nhật danh sách options; nếu reset_selection=True thì đưa về placeholder."""
        self._all_values = list(values)
        if reset_selection:
            self._value.set(self._placeholder)
        if self._listbox and self._popup and self._popup.winfo_exists():
            self._fill_listbox(self._all_values)

    def clear(self):
        """Reset selection về placeholder, đóng popup nếu đang mở."""
        self._value.set(self._placeholder)
        try:
            if self._popup and self._popup.winfo_exists():
                self._popup.withdraw()
        except Exception:
            pass

    def configure(self, **kw):
        if "values" in kw:
            self._all_values = list(kw.pop("values"))
        if "state" in kw:
            self._state = kw.pop("state")
            entry_state = "readonly" if self._state == "normal" else "disabled"
            self._entry.configure(state=entry_state)
            self._btn.configure(state=self._state)
        # CTkFrame không nhận các kw lạ, bỏ qua
        allowed = {k: v for k, v in kw.items()
                   if k in ("fg_color", "border_color", "border_width",
                             "corner_radius", "width", "height")}
        if allowed:
            super().configure(**allowed)

    def set_error(self, flag: bool):
        """Tô đỏ / bỏ đỏ border entry để báo trường bắt buộc chưa chọn."""
        color = ("#EF4444", "#DC2626") if flag else ("#E2E8F0", "#444444")
        self._entry.configure(border_color=color)

    # ── Popup ─────────────────────────────────────────────────────────────────
    def _toggle_popup(self, _event=None):
        if self._state == "disabled":
            return
        if self._popup and self._popup.winfo_exists():
            self._close_popup()
        else:
            self._open_popup()

    def _open_popup(self):
        if self._popup and self._popup.winfo_exists():
            return

        # Đóng popup của combo khác đang mở (tránh 2 dropdown chèn nhau)
        other = _SearchCombo._open_combo
        if other is not None and other is not self:
            try:
                other._close_popup()
            except Exception:
                pass
        _SearchCombo._open_combo = self

        self.update_idletasks()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2
        w = self._popup_width if self._popup_width else max(self.winfo_width(), self._width, 600)

        pop = tk.Toplevel(self)
        pop.overrideredirect(True)
        pop.configure(bg=self._BG_DARK,
                      highlightbackground=self._BORDER,
                      highlightthickness=1)
        pop.geometry(f"{w}x280+{x}+{y}")
        pop.lift()
        pop.attributes("-topmost", True)
        self._popup = pop

        # Search entry
        self._pop_search_var.set("")
        search_entry = tk.Entry(
            pop, textvariable=self._pop_search_var,
            bg=self._BG_ENTRY, fg=self._FG,
            font=self._FONT, insertbackground=self._FG,
            relief="flat", bd=6)
        search_entry.pack(fill=tk.X, padx=4, pady=(4, 2))
        search_entry.focus_set()
        self._pop_search_var.trace_add("write", self._on_search_change)

        # Listbox + scrollbar
        frm = tk.Frame(pop, bg=self._BG_DARK)
        frm.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        vsb = tk.Scrollbar(frm, orient=tk.VERTICAL, bg=self._BG_DARK)
        lb  = tk.Listbox(
            frm, bg=self._BG_DARK, fg=self._FG,
            selectbackground=self._SEL_BG, selectforeground=self._SEL_FG,
            font=self._FONT, relief="flat", bd=0,
            activestyle="none",
            selectborderwidth=0,
            highlightthickness=0,
            yscrollcommand=vsb.set)
        vsb.configure(command=lb.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._listbox = lb

        self._fill_listbox(self._all_values)

        lb.bind("<<ListboxSelect>>", self._on_listbox_select)
        lb.bind("<Return>",          self._on_listbox_select)
        lb.bind("<Escape>",          lambda e: self._close_popup())
        search_entry.bind("<Return>",  self._on_listbox_select)
        search_entry.bind("<Escape>",  lambda e: self._close_popup())
        search_entry.bind("<Down>",    lambda e: lb.focus_set())
        lb.bind("<Up>", lambda e: search_entry.focus_set()
                if lb.curselection() and lb.curselection()[0] == 0 else None)

        # Đóng khi click ngoài hoặc app mất focus (kể cả switch sang app khác)
        pop.bind("<FocusOut>", self._on_popup_focus_out)
        self._poll_focus()

    def _poll_focus(self):
        """Poll 250ms — đóng popup nếu main window mất focus (switch sang app khác)."""
        if not (self._popup and self._popup.winfo_exists()):
            return
        try:
            has_focus = self.winfo_toplevel().focus_displayof() is not None
        except Exception:
            has_focus = True  # nếu không xác định được thì giữ nguyên
        if not has_focus:
            self._close_popup()
            return
        self.after(250, self._poll_focus)

    def _on_search_change(self, *_):
        kw = self._pop_search_var.get().strip().lower()
        if kw:
            filtered = [v for v in self._all_values if kw in v.lower()]
        else:
            filtered = self._all_values
        self._fill_listbox(filtered)

    def _fill_listbox(self, values: list[str]):
        if not (self._listbox and self._popup and self._popup.winfo_exists()):
            return
        self._listbox.delete(0, tk.END)
        for v in values:
            self._listbox.insert(tk.END, v)

    def _on_listbox_select(self, event=None):
        if not self._listbox:
            return
        sel = self._listbox.curselection()
        if not sel:
            if self._listbox.size() > 0:
                idx = 0
            else:
                return
        else:
            idx = sel[0]
        value = self._listbox.get(idx)
        self._value.set(value)
        self._close_popup()
        if self._command:
            self._command(value)

    def _on_popup_focus_out(self, event=None):
        # Delay nhỏ để click vào listbox kịp register trước khi đóng
        self.after(200, self._maybe_close)

    def _maybe_close(self):
        if not (self._popup and self._popup.winfo_exists()):
            return
        try:
            fw = self._popup.focus_get()
        except Exception:
            fw = None
        if fw is None:
            self._close_popup()

    def _close_popup(self):
        if self._popup:
            try:
                self._popup.destroy()
            except Exception:
                pass
            self._popup  = None
            self._listbox = None
        if _SearchCombo._open_combo is self:
            _SearchCombo._open_combo = None


# ── Tooltip ────────────────────────────────────────────────────────────────────
class Tooltip:
    def __init__(self, widget, get_text):
        self._w, self._get, self._tip = widget, get_text, None
        widget.bind("<Enter>",       self._show, add="+")
        widget.bind("<Leave>",       self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _show(self, _=None):
        text = self._get()
        if not text: return
        self._hide()
        dt = THEMES[ctk.get_appearance_mode()]
        x = self._w.winfo_rootx()
        y = self._w.winfo_rooty() + self._w.winfo_height() + 6
        self._tip = tk.Toplevel(self._w)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tip, text=text,
                 bg=dt["bg_card"], fg="#9CDCFE",
                 font=("Segoe UI", 12),
                 relief="flat", padx=10, pady=5,
                 wraplength=700, justify="left").pack()

    def _hide(self, _=None):
        if self._tip:
            try: self._tip.destroy()
            except Exception: pass
            self._tip = None


# ── SheetTable: tksheet với API tối giản của ttk.Treeview ────────────────────
class SheetTable:
    """Bảng dữ liệu dùng tksheet (grid kiểu Excel) nhưng giữ API tối giản của
    ttk.Treeview show='headings' — code đổ dữ liệu cũ (insert / delete /
    heading / column / tag_configure / see / selection_set / set / exists)
    chạy được không đổi. Render được gộp qua after_idle nên insert hàng
    nghìn dòng trong vòng lặp không bị chậm.
    Chỉ khởi tạo khi _HAS_TKSHEET=True; ngược lại dùng ttk.Treeview như cũ."""

    _ANCHOR_MAP = {"w": "w", "e": "e", "center": "center", "c": "center"}

    def __init__(self, master, columns=(), **_kw):
        dt = THEMES[ctk.get_appearance_mode()]
        _sheet_theme = "light" if ctk.get_appearance_mode() == "Light" else "dark"
        self.sheet = tksheet.Sheet(
            master,
            headers=[str(c) for c in columns],
            data=[],
            theme=_sheet_theme,
            show_row_index=False,
            show_top_left=False,
            row_height=30,
            header_height=34,
            show_horizontal_grid=True,
            show_vertical_grid=True,
            font=("Segoe UI", 10, "normal"),
            header_font=("Segoe UI", 10, "bold"),
        )
        self.sheet.enable_bindings(
            "single_select", "drag_select", "column_select", "row_select",
            "column_width_resize", "row_height_resize",
            "arrowkeys", "right_click_popup_menu",
            "rc_insert_row", "rc_delete_row",
            "copy", "cut", "paste", "delete", "undo", "edit_cell",
        )
        # Fine-tune theme — tên option chính xác của tksheet 7.x (short names)
        try:
            self.sheet.set_options(
                table_bg=dt["sheet_table_bg"],
                table_fg=dt["sheet_table_fg"],
                header_bg=dt["sheet_header_bg"],
                header_fg=dt["sheet_header_fg"],
                index_bg=dt["sheet_index_bg"],
                index_fg=dt["sheet_index_fg"],
                top_left_bg=dt["sheet_header_bg"],
                show_horizontal_grid=True,
                show_vertical_grid=True,
                table_grid_fg=dt["sheet_grid"],
                header_grid_fg=dt["sheet_grid"],
                index_grid_fg=dt["sheet_grid"],
                outline_color=dt["sheet_grid"],
                # Scrollbar
                vertical_scroll_troughcolor=dt["scroll_track"],
                vertical_scroll_not_active_bg=dt["scroll_thumb"],
                vertical_scroll_not_active_fg=dt["scroll_thumb"],
                vertical_scroll_active_bg=dt["scroll_thumb_active"],
                vertical_scroll_active_fg=dt["scroll_thumb_active"],
                vertical_scroll_pressed_bg=dt["scroll_thumb_active"],
                vertical_scroll_pressed_fg=dt["scroll_thumb_active"],
                vertical_scroll_bordercolor=dt["scroll_track"],
                vertical_scroll_lightcolor=dt["scroll_track"],
                vertical_scroll_darkcolor=dt["scroll_track"],
                horizontal_scroll_troughcolor=dt["scroll_track"],
                horizontal_scroll_not_active_bg=dt["scroll_thumb"],
                horizontal_scroll_not_active_fg=dt["scroll_thumb"],
                horizontal_scroll_active_bg=dt["scroll_thumb_active"],
                horizontal_scroll_active_fg=dt["scroll_thumb_active"],
                horizontal_scroll_pressed_bg=dt["scroll_thumb_active"],
                horizontal_scroll_pressed_fg=dt["scroll_thumb_active"],
                horizontal_scroll_bordercolor=dt["scroll_track"],
                horizontal_scroll_lightcolor=dt["scroll_track"],
                horizontal_scroll_darkcolor=dt["scroll_track"],
            )
        except Exception:
            pass
        self._columns      = [str(c) for c in columns]
        self._headings     = {c: c for c in self._columns}
        self._col_widths   = {}    # col → width (áp trong _flush)
        self._col_aligns   = {}    # col → 'w'/'e'/'center'
        self._auto_align_all()     # set baseline alignment từ tên cột
        self._applied_cols = list(self._columns)   # bộ cột đã áp vào sheet
        self._style_dirty  = False # có width/align mới chưa áp
        self._tags      = {}     # tag → {'bg':…, 'fg':…}
        self._rows      = []     # thứ tự iid
        self._data      = {}     # iid → list[str]
        self._row_tags  = {}     # iid → tag đầu tiên
        self._center_tag  = None   # tag của các dòng cần căn giữa
        self._italic_tag  = None   # tag của các dòng cần italic (canvas post-render)
        self._italic_font = ("Segoe UI", 9, "italic")
        self._italic_bound = False # đã bind Expose chưa
        self._hidden_cols  = []  # data-index (0-based) cột ẩn — áp sau set_sheet_data
        self._next_iid  = 0
        self._pending   = False

    def center_rows_by_tag(self, tag):
        """Căn giữa mọi dòng mang tag này (vd hàng phụ SQL_Column),
        không ảnh hưởng căn lề cột của các dòng dữ liệu khác."""
        self._center_tag = tag
        self._schedule()

    def italic_rows_by_tag(self, tag, font=("Segoe UI", 9, "italic")):
        """Đặt italic cho các dòng mang tag này bằng cách patch canvas sau render."""
        self._italic_tag  = tag
        self._italic_font = font
        if not self._italic_bound:
            try:
                self.sheet.MT.bind("<Expose>",
                    lambda _e: self.sheet.after_idle(self._apply_italic_canvas))
                self._italic_bound = True
            except Exception:
                pass
        self._schedule()

    def _apply_italic_canvas(self):
        """Post-render: tìm canvas text items ở các dòng italic_tag rồi đổi font."""
        if self._italic_tag is None:
            return
        try:
            MT  = self.sheet.MT
            rps = MT.row_positions      # list[int] — y của từng dòng
            if len(rps) < 2:
                return
            italic_rows = [i for i, iid in enumerate(self._rows)
                           if self._row_tags.get(iid) == self._italic_tag]
            for r in italic_rows:
                if r + 1 >= len(rps):
                    continue
                y1, y2 = rps[r], rps[r + 1]
                for item in MT.find_overlapping(0, y1, 99999, y2):
                    if MT.type(item) == "text":
                        MT.itemconfig(item, font=self._italic_font)
        except Exception:
            pass

    # ── Geometry manager / event delegation ──────────────────────────────────
    def grid(self, **kw):        self.sheet.grid(**kw)
    def pack(self, **kw):        self.sheet.pack(**kw)
    def place(self, **kw):       self.sheet.place(**kw)
    def bind(self, *a, **k):     self.sheet.MT.bind(*a, **k)
    def after(self, *a, **k):    return self.sheet.after(*a, **k)
    def after_cancel(self, *a):  return self.sheet.after_cancel(*a)

    # ── API tương thích Treeview ──────────────────────────────────────────────
    def _auto_align_all(self):
        """Đặt baseline alignment cho tất cả cột theo tên, không ghi đè explicit calls."""
        for col in self._columns:
            if col not in self._col_aligns:
                self._col_aligns[col] = guess_col_align(col)

    def __setitem__(self, key, value):
        if key == "columns":
            self._columns    = [str(c) for c in value]
            self._headings   = {c: c for c in self._columns}
            self._col_widths = {}
            self._col_aligns = {}
            self._auto_align_all()
            self._schedule()

    def __getitem__(self, key):
        if key == "columns":
            return tuple(self._columns)
        raise KeyError(key)

    def heading(self, col, text=None, anchor=None, **_kw):
        if text is not None and col in self._headings:
            self._headings[col] = text
            self._schedule()

    def column(self, col, width=None, anchor=None, **_kw):
        if col not in self._columns:
            return
        if width is not None:
            self._col_widths[col] = int(width)
            self._style_dirty = True
        if anchor is not None:
            self._col_aligns[col] = self._ANCHOR_MAP.get(str(anchor), "w")
            self._style_dirty = True
        self._schedule()

    def tag_configure(self, tag, background=None, foreground=None, **_kw):
        self._tags[tag] = {"bg": background, "fg": foreground}

    def insert(self, _parent, index, values=(), tags=(), **_kw):
        iid = f"r{self._next_iid}"
        self._next_iid += 1
        vals = ["" if v is None else str(v) for v in values]
        if index in (0, "0"):
            self._rows.insert(0, iid)
        else:
            self._rows.append(iid)
        self._data[iid] = vals
        if tags:
            self._row_tags[iid] = (tags[0] if isinstance(tags, (list, tuple))
                                   else str(tags))
        self._schedule()
        return iid

    def delete(self, *iids):
        if not iids:
            return
        drop = set(iids)
        self._rows = [i for i in self._rows if i not in drop]
        for i in drop:
            self._data.pop(i, None)
            self._row_tags.pop(i, None)
        self._schedule()

    def get_children(self, _item=None):
        return tuple(self._rows)

    def exists(self, iid):
        return iid in self._data

    def set(self, iid, col):
        vals = self._data.get(iid)
        if vals is None or col not in self._columns:
            return ""
        idx = self._columns.index(col)
        return vals[idx] if idx < len(vals) else ""

    def item(self, iid, option=None, **_kw):
        vals = self._data.get(iid, [])
        if option == "values":
            return tuple(vals)
        return {"values": tuple(vals)}

    def selection_set(self, iid):
        self._flush()
        if iid in self._data:
            try:
                self.sheet.select_row(self._rows.index(iid))
            except Exception:
                pass

    def selection(self):
        self._flush()
        try:
            return tuple(self._rows[r]
                         for r in sorted(self.sheet.get_selected_rows())
                         if r < len(self._rows))
        except Exception:
            return ()

    def see(self, iid):
        self._flush()
        if iid in self._data:
            try:
                self.sheet.see(row=self._rows.index(iid), column=0)
            except Exception:
                pass

    def focus(self, _iid=None):
        return None

    def config(self, **_kw):
        pass
    configure = config

    def yview(self, *_a):
        pass

    def xview(self, *_a):
        pass

    def load(self, headers: list, data: list):
        """Nạp toàn bộ dữ liệu mới vào SheetTable (reset columns + rows)."""
        self._columns  = [str(c) for c in headers]
        self._headings = {c: c for c in self._columns}
        self._col_widths = {}
        self._col_aligns = {}
        self._auto_align_all()
        self._applied_cols = []
        self._style_dirty  = True
        self._rows     = []
        self._data     = {}
        self._row_tags = {}
        self._next_iid = 0
        for row in data:
            iid = f"r{self._next_iid}"
            self._next_iid += 1
            self._rows.append(iid)
            self._data[iid] = ["" if v is None else str(v) for v in row]
        self._schedule()

    def mark_row(self, idx: int, tag: str):
        """Gán tag highlight cho dòng tại vị trí idx (0-based). Gọi sau load()."""
        if 0 <= idx < len(self._rows):
            self._row_tags[self._rows[idx]] = tag
            self._schedule()

    def hide_columns_by_index(self, cols: list):
        """Ẩn các cột theo data-index (0-based). Áp trong _flush() sau set_sheet_data."""
        self._hidden_cols = list(cols)
        self._schedule()

    def autofit_columns(self, min_w: int = 80, max_w: int = 220):
        """Tự động set width cột theo nội dung thực tế, giới hạn [min_w, max_w].
        Bao gồm cả header và mọi dòng data. Gọi sau load()."""
        try:
            import tkinter.font as tkfont
            _hfont = tkfont.Font(family="Segoe UI", size=10, weight="bold")
            _dfont = tkfont.Font(family="Segoe UI", size=10)
            PAD = 24
            for ci, col in enumerate(self._columns):
                # Đo header
                w = _hfont.measure(self._headings.get(col, col)) + PAD
                # Đo tất cả các ô trong cột
                for iid in self._rows:
                    vals = self._data.get(iid, [])
                    if ci < len(vals):
                        cw = _dfont.measure(str(vals[ci] or "")) + PAD
                        if cw > w:
                            w = cw
                self._col_widths[col] = max(min_w, min(max_w, w))
            self._style_dirty = True
            self._schedule()
        except Exception:
            pass

    # ── Render gộp ────────────────────────────────────────────────────────────
    def _schedule(self):
        if not self._pending:
            self._pending = True
            self.sheet.after_idle(self._flush)

    def _flush(self):
        if not self._pending:
            return
        self._pending = False
        ncols = len(self._columns)
        data  = []
        for iid in self._rows:
            vals = self._data[iid]
            # Pad/cắt cho khớp số cột — tksheet tạo cột theo độ dài row
            if len(vals) < ncols:
                vals = vals + [""] * (ncols - len(vals))
            elif len(vals) > ncols:
                vals = vals[:ncols]
            data.append(vals)
        # Bộ cột thay đổi (VD: Raw Excel gán cột động sau khi đọc file)
        # → phải reset col positions, nếu không tksheet giữ layout cột cũ.
        # So thêm với total_columns() thực tế: bảng được gán cột lúc CHƯA có
        # dữ liệu (VD: tab "Đã xử lý") thì sheet đang có 0 cột dù
        # _applied_cols đã đúng — vẫn phải reset.
        cols_changed = self._applied_cols != self._columns
        try:
            if self.sheet.total_columns() != ncols:
                cols_changed = True
        except Exception:
            pass
        try:
            self.sheet.set_sheet_data(data,
                                      reset_col_positions=cols_changed,
                                      redraw=False)
        except Exception:
            self.sheet.set_sheet_data(data)
        try:
            self.sheet.headers([self._headings.get(c, c)
                                for c in self._columns])
        except Exception:
            pass
        # Lưới an toàn: set_sheet_data/headers đôi khi không co hết cột khi
        # chuyển từ bảng nhiều cột → ít cột (để lại cột trống O,P,Q...).
        # Xoá thẳng các cột dư phía sau.
        if ncols > 0:
            try:
                extra = self.sheet.total_columns() - ncols
                if extra > 0:
                    self.sheet.del_columns(
                        list(range(ncols, ncols + extra)),
                        redraw=False)
            except Exception:
                pass
        if cols_changed or self._style_dirty:
            for col, w in self._col_widths.items():
                if col in self._columns:
                    try:
                        self.sheet.column_width(
                            column=self._columns.index(col), width=w)
                    except Exception:
                        pass
            for col, al in self._col_aligns.items():
                if col in self._columns:
                    idx = self._columns.index(col)
                    try:
                        self.sheet.align_columns(columns=[idx], align=al)
                    except Exception:
                        pass
            self._style_dirty = False
        # Header luôn căn giữa, áp mỗi flush để chắc không bị reset
        _n = len(self._columns)
        if _n > 0:
            try:
                self.sheet.align_header(
                    columns=list(range(_n)), align='center', redraw=False
                )
            except Exception:
                pass
        self._applied_cols = list(self._columns)
        try:
            self.sheet.dehighlight_all()
        except Exception:
            pass
        by_tag = {}
        for idx, iid in enumerate(self._rows):
            t = self._row_tags.get(iid)
            if t in self._tags:
                by_tag.setdefault(t, []).append(idx)
        for t, idxs in by_tag.items():
            cfg = self._tags[t]
            if not (cfg.get("bg") or cfg.get("fg")):
                continue
            try:
                self.sheet.highlight_rows(rows=idxs, bg=cfg.get("bg"),
                                          fg=cfg.get("fg"), redraw=False)
            except Exception:
                pass
        # Căn giữa các dòng mang center-tag (vd hàng phụ SQL_Column) — ghi đè
        # căn lề cột, không đụng các dòng dữ liệu khác
        if self._center_tag is not None:
            c_idx = [i for i, iid in enumerate(self._rows)
                     if self._row_tags.get(iid) == self._center_tag]
            if c_idx:
                try:
                    self.sheet.align_rows(rows=c_idx, align='center', redraw=False)
                except Exception:
                    pass
        # Dòng sub-header (tag "sql_names"): readonly + nhỏ hơn + italic
        _sub_hdr = [i for i, iid in enumerate(self._rows)
                    if self._row_tags.get(iid) == "sql_names"]
        if _sub_hdr:
            try:
                self.sheet.readonly_rows(rows=_sub_hdr, readonly=True, redraw=False)
            except Exception:
                pass
            try:
                for r in _sub_hdr:
                    self.sheet.row_height(row=r, height=26, redraw=False)
            except Exception:
                pass
            try:
                self.sheet.align_rows(rows=_sub_hdr, align='center', redraw=False)
            except Exception:
                pass
            # Italic font — bỏ qua nếu phiên bản tksheet không hỗ trợ font param
            try:
                self.sheet.highlight_rows(
                    rows=_sub_hdr, font=("Segoe UI", 9, "italic"), redraw=False)
            except (TypeError, AttributeError, Exception):
                pass
        try:
            self.sheet.refresh()
        except Exception:
            pass
        # Ẩn cột sau refresh — phải gọi sau set_sheet_data/refresh để không bị reset
        if self._hidden_cols:
            try:
                self.sheet.hide_columns(columns=self._hidden_cols, data_indexes=True)
            except Exception:
                pass
        if self._italic_tag is not None:
            self.sheet.after_idle(self._apply_italic_canvas)

    def apply_theme(self, t: dict):
        """Cập nhật màu sắc tksheet theo palette theme t (Dark/Light)."""
        _st  = t.get("scroll_track",        "#252526")
        _sfg = t.get("scroll_thumb",        "#3E3E42")
        _sac = t.get("scroll_thumb_active", "#606066")
        try:
            self.sheet.set_options(
                table_bg=t["sheet_table_bg"],
                table_fg=t["sheet_table_fg"],
                header_bg=t["sheet_header_bg"],
                header_fg=t["sheet_header_fg"],
                index_bg=t["sheet_index_bg"],
                index_fg=t["sheet_index_fg"],
                top_left_bg=t["sheet_header_bg"],
                show_horizontal_grid=True,
                show_vertical_grid=True,
                table_grid_fg=t["sheet_grid"],
                header_grid_fg=t["sheet_grid"],
                index_grid_fg=t["sheet_grid"],
                outline_color=t["sheet_grid"],
                # Scrollbar
                vertical_scroll_troughcolor=_st,
                vertical_scroll_not_active_bg=_sfg,
                vertical_scroll_not_active_fg=_sfg,
                vertical_scroll_active_bg=_sac,
                vertical_scroll_active_fg=_sac,
                vertical_scroll_pressed_bg=_sac,
                vertical_scroll_pressed_fg=_sac,
                vertical_scroll_bordercolor=_st,
                vertical_scroll_lightcolor=_st,
                vertical_scroll_darkcolor=_st,
                horizontal_scroll_troughcolor=_st,
                horizontal_scroll_not_active_bg=_sfg,
                horizontal_scroll_not_active_fg=_sfg,
                horizontal_scroll_active_bg=_sac,
                horizontal_scroll_active_fg=_sac,
                horizontal_scroll_pressed_bg=_sac,
                horizontal_scroll_pressed_fg=_sac,
                horizontal_scroll_bordercolor=_st,
                horizontal_scroll_lightcolor=_st,
                horizontal_scroll_darkcolor=_st,
            )
            self.sheet.redraw()
        except Exception:
            pass

    def rehighlight(self):
        """Re-apply row highlights after tag_configure(). Nhẹ hơn _flush() vì không reload data."""
        if not self._rows:
            return
        try:
            self.sheet.dehighlight_all()
            by_tag: dict[str, list[int]] = {}
            for idx, iid in enumerate(self._rows):
                tg = self._row_tags.get(iid)
                if tg in self._tags:
                    by_tag.setdefault(tg, []).append(idx)
            for tg, idxs in by_tag.items():
                cfg = self._tags[tg]
                if cfg.get("bg") or cfg.get("fg"):
                    self.sheet.highlight_rows(rows=idxs,
                                              bg=cfg.get("bg"),
                                              fg=cfg.get("fg"),
                                              redraw=False)
            self.sheet.redraw()
        except Exception:
            pass

