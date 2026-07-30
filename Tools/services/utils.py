"""
services/utils.py — Hằng số, path, helper thuần Python (không UI).
"""
import re, os, sys, unicodedata, datetime, math

APP_VERSION = "2.1"

# Tên 3 tab chính — nguồn duy nhất, dùng ở mọi chỗ thay vì literal string /
# substring match. Trước đây _on_tab_changed() check `"THDM" in tab_name`,
# khi đổi TAB_THDM sang "Tổng hợp định mức" thì chuỗi không còn chứa "THDM"
# nữa nên auto-reload khi switch tab bị vô hiệu âm thầm — đổi tên tab từ giờ
# chỉ cần sửa đúng 1 dòng dưới đây.
TAB_IMPORT  = "📥  Import BOM"
TAB_THDM    = "📊  Tổng hợp định mức"
TAB_CATALOG = "🌳  Danh mục vật tư"


def get_base_dir():
    if getattr(sys, 'frozen', False):          # Đang chạy là .exe (PyInstaller)
        return os.path.dirname(sys.executable)
    # __file__ = .../Tools/services/utils.py → cần trả về .../Tools/
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BASE_DIR = get_base_dir()


def resource_path(relative_path: str) -> str:
    """Trả về đường dẫn tuyệt đối tới tài nguyên CHỈ-ĐỌC được bundle vào exe.
    - Khi chạy từ .exe (PyInstaller): tìm trong sys._MEIPASS (_internal/).
    - Khi chạy dev: tìm trong BASE_DIR (thư mục Tools/).
    Không dùng cho config files (db_config, CK_Mapping) — dùng BASE_DIR cho đó."""
    try:
        base = sys._MEIPASS           # type: ignore[attr-defined]
    except AttributeError:
        base = BASE_DIR
    return os.path.join(base, relative_path)
CONFIG_DIR     = os.path.join(BASE_DIR, "config")
MAPPING_FILE   = os.path.join(CONFIG_DIR, "CK_Mapping_v5.xlsx")
DB_CONFIG_FILE = os.path.join(CONFIG_DIR, "db_config.json")
SETTINGS_FILE  = os.path.join(CONFIG_DIR, "settings.json")

# UserId mặc định dùng cho CreatedBy khi nhân viên được chọn không có tài
# khoản tương ứng trong B00UserList (EmployeeId không map được sang Id user).
DEFAULT_CREATOR_USER_ID = 823

os.makedirs(CONFIG_DIR, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# 2. MAPPING LOADER — đọc CK_Mapping_v2.xlsx (format mới, mỗi sheet = 1 section)
# ─────────────────────────────────────────────────────────────────────────────

def _norm_vn(s):
    """Chuẩn hóa tiếng Việt: bỏ dấu, lowercase, giữ chữ/số.
    'đ' (U+0111) không decompose bằng NFKD — xử lý riêng."""
    s = str(s).lower().strip()
    s = s.replace('đ', 'd')
    nfkd = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r'[^a-z0-9]', '', s)


def guess_col_align(col_name: str) -> str:
    """Suy đoán căn lề tksheet/Treeview từ tên cột:
      'e'      → số lượng / kích thước / giá (căn phải)
      'center' → mã / STT / loại / flag / DVT (căn giữa)
      'w'      → tên / mô tả / ghi chú (căn trái, mặc định)
    """
    raw = str(col_name).strip()
    # Ký tự đặc biệt (checkbox, tick) → căn giữa
    if not raw or raw in ('✓', '□', '☐', '■', '-'):
        return 'center'
    # Bỏ dấu + lowercase + chỉ giữ [a-z0-9]
    _s = raw.lower().replace('đ', 'd')
    _nfkd = unicodedata.normalize('NFKD', _s)
    n = re.sub(r'[^a-z0-9]', '',
               ''.join(c for c in _nfkd if not unicodedata.combining(c)))
    if not n:
        return 'center'

    # ── PHẢI: số lượng / kích thước / giá tiền ──────────────────────────────
    _R_EXACT = {
        'day', 'rong', 'dai',                           # DÀY RỘNG DÀI
        'sl', 'slg', 'soluong', 'soluongle',            # SLg, Số lượng
        'khoiluong', 'trongluong', 'kl', 'tl',         # Khối lượng
        'gia', 'dongia', 'tong', 'tongdong', 'tongso', # Giá, Tổng
        'tongcong', 'thanhtien', 'tyle', 'phantram',    # Thành tiền, %
        'qty', 'quantity', 'weight', 'width', 'height', # English
        'length', 'thickness', 'depth', 'price',
        'amount', 'total', 'subtotal', 'count',
        'size', 'area', 'volume', 'mass',
    }
    _R_SUB = (
        'soluong', 'khoiluong', 'trongluong', 'dongia', 'thanhtien',
        'quantity', 'weight', 'width', 'height', 'length', 'thickness',
        'price', 'amount',
    )
    if n in _R_EXACT:
        return 'e'
    if n.startswith('sl') and not n.startswith('sql'):   # sl... nhưng không phải sql...
        return 'e'
    for _sub in _R_SUB:
        if _sub in n:
            return 'e'

    # ── GIỮA: mã / STT / loại / flag / đơn vị ──────────────────────────────
    _C_EXACT = {
        'stt', 'mucso', 'muc',                          # STT, Mục số
        'dvt', 'donviti', 'unit',                       # DVT
        'version', 'ver',                               # Version
        'loai', 'kieudl', 'kieu',                      # Kiểu DL
        'batbuoc', 'kieulookup',                        # mapping meta
        'trangthai', 'sheet', 'check',                  # log/status
        'type', 'status', 'flag', 'active', 'required',
    }
    _C_PREFIX = ('ma',)           # Mã... (mã SP, mã BOM, mã VT...)
    _C_SUFFIX = ('id', 'no', 'code', 'ma', 'type', 'status')

    if n in _C_EXACT:
        return 'center'
    for _px in _C_PREFIX:
        if n.startswith(_px) and 2 <= len(n) <= 15:
            return 'center'
    for _sx in _C_SUFFIX:
        if n != _sx and n.endswith(_sx) and len(n) > len(_sx):
            return 'center'

    return 'w'

# Cột mapping chuẩn (14 cột, không có STT)
_MAPPING_ENG_COLS = [
    "SQL_Column", "Ten_Excel", "Kieu_DL", "Do_dai",
    "Bat_buoc", "Mac_dinh", "Nguon_DL",
    "Bang_Master", "Dieu_kien_Master", "Kieu_Lookup",
    "Truong_So_Sanh", "Truong_Lay_Ve", "Ghi_chu",
    "Fill_Forward",
]
# HEADER/DETAIL sheet có thêm cột Section ở đầu (15 cột)
_SECTION_ENG_COLS = ["Section"] + _MAPPING_ENG_COLS


def _nan_str(v, default=''):
    """Chuyển pandas NaN / None thành chuỗi rỗng."""
    import math as _math
    if v is None:
        return default
    if isinstance(v, float) and _math.isnan(v):
        return default
    s = str(v).strip()
    return default if s.lower() == 'nan' else s


# ─────────────────────────────────────────────────────────────────────────────
# 5. GUI
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# THEME PALETTES — nguồn duy nhất cho cả Dark và Light mode
# ─────────────────────────────────────────────────────────────────────────────
THEMES = {
    "Dark": {
        # App backgrounds
        "bg_main":          "#1E1E1E",
        "bg_card":          "#252526",
        "bg_panel":         "#2D2D2D",
        "border":           "#333333",
        "text_main":        "#E8E8E8",
        "text_muted":       "#8A8A8A",
        "accent":           "#4FC3F7",
        # tksheet
        "sheet_header_bg":  "#1B2A4A",
        "sheet_header_fg":  "#FFFFFF",
        "sheet_table_bg":   "#1E1E1E",
        "sheet_table_fg":   "#E8E8E8",
        "sheet_grid":       "#888888",
        "sheet_index_bg":   "#252526",
        "sheet_index_fg":   "#E8E8E8",
        # Treeview
        "tv_text":          "#CCCCCC",
        "tv_sel":           "#37373D",
        "tv_sep":           "#3C3C3C",
        "tv_heading_bg":    "#1B2A4A",
        "tv_heading_fg":    "#FFFFFF",
        "tv_heading_act":   "#253A60",
        # SearchCombo popup
        "popup_bg":         "#252526",
        "popup_border":     "#3D3D3D",
        "popup_entry_bg":   "#3C3C3C",
        "sel_bg":           "#0066CC",
        # Catalog tree tag fg
        "cat_leaf":         "#CCCCCC",
        "cat_inactive":     "#5A7070",
        "cat_orphan":       "#7A7A7A",
        # Log tree tag fg
        "log_ok":           "#4EC9B0",
        "log_warn":         "#D7BA7D",
        "log_err":          "#F48771",
        # Row highlight tags (SheetTable + Treeview)
        "field_row_bg":     "#223355",
        "field_row_fg":     "#A0C4FF",
        "row_normal_fg":    "#E8E8E8",
        "row_alt_fg":       "#8A8A8A",
        "badge_err_bg":     "#3B1C1C",
        "badge_err_fg":     "#FF8F8F",
        "badge_warn_bg":    "#3B2E1C",
        "badge_warn_fg":    "#FFD180",
        # tksheet scrollbar
        "scroll_track":       "#252526",
        "scroll_thumb":       "#3E3E42",
        "scroll_thumb_active":"#606066",
    },
    "Light": {
        # App backgrounds
        "bg_main":          "#F8FAFC",
        "bg_card":          "#FFFFFF",
        "bg_panel":         "#F1F5F9",
        "border":           "#E2E8F0",
        "text_main":        "#0F172A",
        "text_muted":       "#64748B",
        "accent":           "#0066CC",
        # tksheet
        "sheet_header_bg":  "#E2E8F0",
        "sheet_header_fg":  "#0F172A",
        "sheet_table_bg":   "#FFFFFF",
        "sheet_table_fg":   "#0F172A",
        "sheet_grid":       "#CBD5E1",
        "sheet_index_bg":   "#F1F5F9",
        "sheet_index_fg":   "#0F172A",
        # Treeview
        "tv_text":          "#1E293B",
        "tv_sel":           "#BFDBFE",
        "tv_sep":           "#CBD5E1",
        "tv_heading_bg":    "#E2E8F0",
        "tv_heading_fg":    "#0F172A",
        "tv_heading_act":   "#D1D5DB",
        # SearchCombo popup
        "popup_bg":         "#FFFFFF",
        "popup_border":     "#CBD5E1",
        "popup_entry_bg":   "#F8FAFC",
        "sel_bg":           "#0066CC",
        # Catalog tree tag fg
        "cat_leaf":         "#1E293B",
        "cat_inactive":     "#64748B",
        "cat_orphan":       "#94A3B8",
        # Log tree tag fg
        "log_ok":           "#059669",
        "log_warn":         "#D97706",
        "log_err":          "#DC2626",
        # Row highlight tags (SheetTable + Treeview)
        "field_row_bg":     "#EFF6FF",
        "field_row_fg":     "#1E3A8A",
        "row_normal_fg":    "#0F172A",
        "row_alt_fg":       "#334155",
        "badge_err_bg":     "#FEE2E2",
        "badge_err_fg":     "#DC2626",
        "badge_warn_bg":    "#FEF3C7",
        "badge_warn_fg":    "#D97706",
        # tksheet scrollbar
        "scroll_track":       "#F1F5F9",
        "scroll_thumb":       "#CBD5E1",
        "scroll_thumb_active":"#94A3B8",
    },
}

C = {
    "bg":     "#1E1E1E",
    "panel":  "#252526",
    "border": "#333333",
    "accent": "#4FC3F7",
    "green":  "#4CAF50",
    "yellow": "#FFB74D",
    "red":    "#EF5350",
    "text":   "#E8E8E8",
    "muted":  "#8A8A8A",
    "white":  "#FFFFFF",
    # Pastel badge colors cho dòng lỗi/cảnh báo
    "badge_err_bg":  "#3B1C1C",
    "badge_err_fg":  "#FF8F8F",
    "badge_warn_bg": "#3B2E1C",
    "badge_warn_fg": "#FFD180",
    # Header style — dùng cho mọi bảng (tksheet + Treeview)
    "header_bg":    "#1B2A4A",   # nền header nổi bật
    "header_fg":    "#FFFFFF",   # chữ header trắng
    "header_grid":  "#3D4148",   # lưới header
    # Dòng tên field SQL (dòng 0 tiếng Anh như ParentBizDocId, Quantity...)
    "field_row_bg": "#223355",
    "field_row_fg": "#A0C4FF",
}

