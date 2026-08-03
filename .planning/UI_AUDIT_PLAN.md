# UI/UX Audit Plan — BOHO IMPORT v2.1

**Ngày audit:** 2026-08-03  
**Công cụ:** Manual review · Grep · Agent code scan  
**Files chính:** `Tools/views/main_window.py` (9,678 dòng) · `Tools/views/widgets.py`

---

## Tóm tắt điểm mạnh (không cần sửa)

| Hạng mục | Đánh giá |
|----------|----------|
| Threading cho tác vụ lâu | ✅ 30+ lệnh `threading.Thread(daemon=True)` — không chặn UI |
| Button disable khi xử lý | ✅ Hơn 20 button/combo bị disable đúng cách |
| Loading overlay | ✅ Startup fullscreen + modal popup per-operation |
| Log panel | ✅ Bottom panel đầy đủ: file log + DB log + Treeview |
| Success feedback | ✅ Triple feedback: status bar + modal + log entry |
| Empty state | ✅ Dedicated empty-state UI trên cả 2 tab chính |

---

## Nhóm U — UX Issues cần sửa (10 issues)

### U1 · Raw traceback hiển thị thẳng cho user — 🔴 Cao

**Vấn đề:** `main_window.py:9428` lưu `_tb.format_exc()` vào `result['error']`. Sau đó tại line `9451`:
```python
self._show_msg("Lỗi chuẩn bị dữ liệu", result['error'][:800], 'error')
```
User nhìn thấy stack trace Python thô (`File "...", line N, in _parse_section_excel_rows`...) — không thân thiện, không có nghĩa gì với người dùng cuối.

**Fix:** Tách message kỹ thuật ra khỏi message user:
```python
# Hiện cho user
user_msg = "Lỗi khi đọc dữ liệu từ file Excel. Vui lòng kiểm tra lại file."
self._show_msg("Lỗi chuẩn bị dữ liệu", user_msg, 'error')
# Ghi kỹ thuật vào log panel
self._log("error", "prescan", result['error'])
```

---

### U2 · 4 dialog legacy `messagebox` còn sót — 🔴 Cao

**Vấn đề:** Ứng dụng đã có `_show_msg()` / `_ask_msg()` riêng (CTk-styled, dark mode). Nhưng 4 chỗ vẫn gọi native `tkinter.messagebox` → popup xuất hiện với style Windows cũ, không match dark theme.

**Vị trí:**
- `main_window.py:5599` — `messagebox.askyesno("Áp dụng hàng loạt?", ...)`
- `main_window.py:6421` — `mb.askyesno("Xác nhận tạo THDM", ...)`
- `main_window.py:7014` — `messagebox.showwarning("Sai mật khẩu", ...)`
- `main_window.py:9479` — `_mb.askyesno(...)`

**Fix:** Thay tất cả bằng `self._ask_msg(...)` hoặc `self._show_msg(...)` tương ứng.

---

### U3 · Thiếu binding `<Escape>` / `<Return>` ở các dialog chính — 🔴 Cao

**Vấn đề:** Các dialog quan trọng không xử lý phím tắt — người dùng phải dùng chuột.

| Dialog | Thiếu binding | File:Line |
|--------|---------------|-----------|
| `_ask_excel_password()` | `<Escape>` để hủy | `bom_parser.py:551-611` |
| `_show_msg()` (thông báo) | `<Return>` / `<Escape>` để đóng | `main_window.py:8981` |
| `_ask_msg()` (yes/no) | `<Return>` = Có, `<Escape>` = Không | `main_window.py:9050` |
| `_show_export_success()` | `<Return>` / `<Escape>` để đóng | `main_window.py:9014` |
| Fuzzy mapping dialog | `<Escape>` để hủy | `main_window.py:7826` |

**Fix:** Thêm binding vào mỗi dialog:
```python
# Pattern chuẩn cho tất cả
dialog.bind("<Escape>", lambda e: on_cancel())
dialog.bind("<Return>", lambda e: on_ok())
dialog.focus_set()
```

---

### U4 · Màu hardcode ~100+ chỗ, không qua theme system — 🟠 Trung bình

**Vấn đề:** THEMES dict trong `utils.py` không được dùng nhất quán. Rất nhiều màu hardcode trực tiếp → khi switch Dark/Light, các widget này KHÔNG đổi màu.

**Hotspot:**
- `bom_parser.py:569-607` — password dialog: `"#1e1e1e"`, `"#334155"`, `"#2563eb"` hardcode
- `main_window.py:561-690` — startup overlay: 12 màu hex thô
- `main_window.py:900-971` — import tab body: `"#251515"`, `"#7B1A1A"`, `"#FFCCCC"` không có trong THEMES

**Fix đề xuất:**
1. Ưu tiên cao: Bổ sung các màu thiếu vào THEMES (`badge_err_bg_dark`, `danger_bg`...) và thay thế ở 3 hotspot trên trước.
2. Ưu tiên thấp hơn: Dần dần migrate toàn bộ 100+ hex literal.

---

### U5 · Mixed `tk.Button` và `ctk.CTkButton` trong dialog — 🟠 Trung bình

**Vấn đề:** Một số dialog dùng `tk.Button` thô (native Windows look) lẫn với `ctk.CTkButton`:
- `bom_parser.py:602-607` — password dialog dùng `tk.Button` với `bg="#334155"`
- `main_window.py:7793-7796` — fuzzy picker dùng `tk.Button(bg=C["green"])`
- Các dialog `_show_msg`, `_ask_msg` dùng `ctk.CTkButton` ✓

Visual mismatch: khi dark mode, `tk.Button` vẫn có nền Windows-gray, trông không đồng nhất.

**Fix:** Thay toàn bộ `tk.Button` trong dialog code bằng `ctk.CTkButton`.

---

### U6 · Không có font constants — 10 kích thước hardcode rải rác — 🟠 Trung bình

**Vấn đề:** Không có hệ thống type scale. 10 kích thước font khác nhau (9, 10, 11, 12, 13, 15, 18, 24, 42, 48) được hardcode tại 60+ vị trí. Size 11 xuất hiện tại 30+ chỗ dưới dạng literal `("Segoe UI", 11)`.

**Fix:** Định nghĩa type scale ở đầu `main_window.py` hoặc trong `utils.py`:
```python
FONT_BODY   = ("Segoe UI", 11)
FONT_SMALL  = ("Segoe UI", 10)
FONT_LABEL  = ("Segoe UI", 11, "bold")
FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_ICON   = ("Segoe UI", 42)
```
Sau đó replace dần.

---

### U7 · Padding/margin không nhất quán — 🟠 Trung bình

**Vấn đề:** 14 giá trị `padx` khác nhau (0–28) và 17 giá trị `pady` khác nhau (0–26). Không có spacing tokens.

**Mẫu mâu thuẫn điển hình:**
- Toolbar button: `padx=(0, 4)` tại line 841 vs `padx=(0, 6)` tại line 874 vs `padx=(0, 8)` tại line 1402
- Section header: `padx=12` tại 1434, `padx=10` tại 1301, `padx=8` tại 1372

**Fix:** Định nghĩa spacing constants:
```python
PAD_XS = 4   # tight gaps
PAD_SM = 8   # default element spacing
PAD_MD = 12  # section padding
PAD_LG = 16  # card/panel margin
```
Apply dần, bắt đầu từ toolbar và section headers (dễ nhìn thấy nhất).

---

### U8 · Progress không có percentage — chỉ spinner — 🟡 Thấp

**Vấn đề:** `_make_loading_popup()` chỉ dùng `mode="indeterminate"` (vòng tròn chạy mãi). Khi import 1000 dòng BOM, user không biết đã xong bao nhiêu %. `_update_loading_msg()` cập nhật text phase, giúp một phần nhưng không có số.

**Fix đề xuất:** Thêm determinate progress bar với `pb.set(done/total)` và update từ worker thread qua `self.after(0, ...)`. Ưu tiên cho BOM insert (đã biết tổng số dòng trước khi chạy).

---

### U9 · Không có toast notification — feedback tồn tại quá lâu hoặc quá ít — 🟡 Thấp

**Vấn đề:** Không có self-dismissing toast. Sau các action nhỏ (copy, export Excel, toggle setting), feedback chỉ là status bar label tĩnh — người dùng dễ bỏ qua.

**Fix:** Thêm hàm `_toast(msg, duration=2500, kind="info")` dùng `self.after(duration, label.destroy)`. Dùng cho các thao tác nhỏ không cần modal.

---

### U10 · Không có explicit re-entrancy guard ở `_start_import` — 🟡 Thấp

**Vấn đề:** `btn_import` không bị disable trước khi gọi `_start_import()` (line 9126). Nếu user click nhanh 2 lần trước khi loading popup xuất hiện (line 9179), có thể khởi động 2 lần import.

**Fix:** Thêm flag hoặc disable button ngay ở đầu `_start_import`:
```python
def _start_import(self):
    self.btn_import.configure(state="disabled")
    try:
        ...
    finally:
        self.btn_import.configure(state="normal")
```

---

## Thứ tự thực hiện đề xuất

```
PHASE 1 — Critical UX fixes (ảnh hưởng ngay đến user)
  ├─ U1: Ẩn traceback khỏi error dialog → log panel
  ├─ U2: Thay 4 legacy messagebox → _show_msg / _ask_msg
  └─ U3: Thêm <Escape>/<Return> vào 5 dialog

PHASE 2 — Visual consistency (ảnh hưởng dark/light mode)
  ├─ U4: Migrate 3 hotspot màu hardcode vào THEMES
  └─ U5: Thay tk.Button → ctk.CTkButton trong dialog

PHASE 3 — Code quality (codebase health)
  ├─ U6: Định nghĩa font constants
  └─ U7: Định nghĩa spacing constants + apply toolbar/section

PHASE 4 — Enhancement (nice-to-have)
  ├─ U8: Progress bar determinate cho BOM insert
  ├─ U9: Toast notification cho action nhỏ
  └─ U10: Re-entrancy guard _start_import
```

**Ước tính:** Phase 1 ~1-2 giờ · Phase 2 ~2-3 giờ · Phase 3 ~2 giờ · Phase 4 ~3 giờ

---

*Chờ duyệt trước khi thực hiện bất kỳ phase nào.*
