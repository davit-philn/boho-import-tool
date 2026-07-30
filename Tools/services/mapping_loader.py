"""
services/mapping_loader.py — Đọc CK_Mapping_v5.xlsx và build reverse-map.
"""
import pandas as pd
import re, os, sys

from services.utils import (
    _norm_vn, _nan_str,
    _MAPPING_ENG_COLS, _SECTION_ENG_COLS,
    MAPPING_FILE, BASE_DIR,
)


# ─── Helper readers cho CK_Mapping_v5 (7 sheets) ───────────────────────────

def _load_config(xl):
    """Đọc _CONFIG → {section: {label, view_insert, sheet_contains, ...}}"""
    if '_CONFIG' not in xl.sheet_names:
        return {}
    df = pd.read_excel(xl, sheet_name='_CONFIG', header=None, skiprows=3)
    COLS = ['Section', 'Label', 'View_Insert', 'SheetNameContains',
            'SheetNameExclude', 'DataStartRow', 'ParentSection',
            'ExpandMuc', 'RowFilter']
    for i, c in enumerate(COLS):
        if i < len(df.columns):
            df.rename(columns={df.columns[i]: c}, inplace=True)
    result = {}
    for _, row in df.iterrows():
        sec = _nan_str(row.get('Section'))
        if not sec:
            continue
        _em = _nan_str(row.get('ExpandMuc'))   # NaN/None/empty → ''
        try:
            _expand_muc = int(float(_em)) != 0 if _em else False
        except (ValueError, TypeError):
            _expand_muc = False
        result[sec] = {
            'label':          _nan_str(row.get('Label')),
            'view_insert':    _nan_str(row.get('View_Insert')),
            'sheet_contains': _nan_str(row.get('SheetNameContains')),
            'sheet_exclude':  _nan_str(row.get('SheetNameExclude')),
            'data_start_row': _nan_str(row.get('DataStartRow')),
            'parent_section': _nan_str(row.get('ParentSection')),
            'expand_muc':     _expand_muc,
            'row_filter':     _nan_str(row.get('RowFilter')),
        }
    return result


def _load_config_rows(xl, sheet_name, col_names):
    """Generic reader: sheet dạng config (skiprows=3), trả về list[dict]."""
    if sheet_name not in xl.sheet_names:
        return []
    df = pd.read_excel(xl, sheet_name=sheet_name, header=None, skiprows=3)
    for i, c in enumerate(col_names):
        if i < len(df.columns):
            df.rename(columns={df.columns[i]: c}, inplace=True)
    records = []
    for _, row in df.iterrows():
        rec = {c.lower(): _nan_str(row.get(c)) for c in col_names if c in df.columns}
        if any(rec.values()):
            records.append(rec)
    return records


def _load_section_rows(xl, sheet_name, section):
    """Đọc HEADER hoặc DETAIL sheet, filter theo cột Section."""
    if sheet_name not in xl.sheet_names:
        return []
    df = pd.read_excel(xl, sheet_name=sheet_name, header=None, skiprows=3)
    for i, c in enumerate(_SECTION_ENG_COLS):
        if i < len(df.columns):
            df.rename(columns={df.columns[i]: c}, inplace=True)
    if 'Section' not in df.columns:
        return []
    df = df[df['Section'].astype(str).str.strip() == section]
    records = []
    for _, row in df.iterrows():
        sql_col = _nan_str(row.get('SQL_Column'))
        if not sql_col or sql_col == 'SQL_Column':
            continue
        records.append({
            "sql_col":          sql_col,
            "ten_excel":        _nan_str(row.get("Ten_Excel")),
            "kieu_dl":          _nan_str(row.get("Kieu_DL")),
            "do_dai":           _nan_str(row.get("Do_dai")),
            "bat_buoc":         _nan_str(row.get("Bat_buoc")),
            "mac_dinh":         _nan_str(row.get("Mac_dinh")),
            "nguon_dl":         _nan_str(row.get("Nguon_DL"), "Excel"),
            "bang_master":      _nan_str(row.get("Bang_Master")),
            "dieu_kien_master": _nan_str(row.get("Dieu_kien_Master")),
            "kieu_lookup":      _nan_str(row.get("Kieu_Lookup")),
            "truong_so_sanh":   _nan_str(row.get("Truong_So_Sanh")),
            "truong_lay_ve":    _nan_str(row.get("Truong_Lay_Ve")),
            "ghi_chu":          _nan_str(row.get("Ghi_chu")),
            "fill_forward":     "1" if str(_nan_str(row.get("Fill_Forward"))).split('.')[0] == '1' else "",
        })
    return records


def load_mapping(mapping_path=MAPPING_FILE):
    """
    Đọc CK_Mapping_v5.xlsx — 7 sheets mới.
    Trả về dict với keys:
      _CONFIG  → {section: metadata}
      HEADER, BOM2, BOM3, BOM4, THDM_HEADER, THDM_THVT → list[record]
      _SP_CONFIG / SP_CONFIG (alias)  → list[record]
      SP_HOOK, VALIDATORS             → list[record]
    Trả {} nếu file không tồn tại.
    """
    if not os.path.exists(mapping_path):
        return {}
    result = {}
    try:
        xl = pd.ExcelFile(mapping_path, engine='openpyxl')

        # 1) _CONFIG — registry tất cả sections
        config = _load_config(xl)
        result['_CONFIG'] = config

        # 2) HEADER/DETAIL sections — derive từ _CONFIG
        for section, cfg in config.items():
            sheet_name = 'HEADER' if not cfg.get('parent_section') else 'DETAIL'
            result[section] = _load_section_rows(xl, sheet_name, section)

        # 3) _SP_CONFIG (gộp BOM + THDM, section rõ ràng)
        SP_COLS = ['Section', 'SQL_Column', 'SP_Name', 'Params',
                   'OutputFields', 'Fallback', 'IsActive']
        sp_rows = _load_config_rows(xl, '_SP_CONFIG', SP_COLS)
        for rec in sp_rows:  # normalize fallback '1.0' → '1'
            fb = rec.get('fallback', '')
            try:
                f = float(fb)
                if f == int(f):
                    rec['fallback'] = str(int(f))
            except (ValueError, TypeError):
                pass
        result['_SP_CONFIG'] = sp_rows
        result['SP_CONFIG']  = sp_rows   # backward compat

        # 4) SP_HOOK
        HOOK_COLS = ['Section', 'Event', 'SP_Name', 'Condition', 'OutputFields',
                     'Params', 'IsActive', 'XmlParam', 'XmlTag', 'XmlFields']
        result['SP_HOOK'] = [
            r for r in _load_config_rows(xl, 'SP_HOOK', HOOK_COLS)
            if r.get('isactive') == '1'
        ]

        # 5) VALIDATORS
        VAL_COLS = ['Section', 'ValidatorName', 'SQL', 'Params',
                    'WarningMessage', 'WarnOnly', 'IsActive']
        result['VALIDATORS'] = [
            r for r in _load_config_rows(xl, 'VALIDATORS', VAL_COLS)
            if r.get('validatorname')
        ]

    except Exception as e:
        print(f"[Mapping] Lỗi đọc file mapping: {e}")
    return result


# ── VALIDATORS compat: old code dùng VAL_COLS không có Section/WarnOnly ────
# Không cần thay đổi — các record đã có key 'section' và 'warnonly' thêm vào
# downstream code đọc 'validatorname','sql','params','warningmessage','isactive'
# sẽ tự bỏ qua key lạ. WarnOnly check thêm khi cần.


# ── CATALOG_FILTER — đọc trực tiếp qua openpyxl (giữ nguyên, chỉ đổi path) ─
# _catalog_read_filter_ids dùng MAPPING_FILE (đã update line 272)





def build_reverse_map(mapping, section='BOM2'):
    """
    {norm_ten_excel: sql_col} — chỉ với các trường Nguon_DL=Excel.
    Lưu cả key đầy đủ lẫn key bỏ phần trong ngoặc.
    """
    reverse = {}
    for r in mapping.get(section, []):
        if r["nguon_dl"] != "Excel":
            continue
        ten = r["ten_excel"]
        sql = r["sql_col"]
        norm_full  = _norm_vn(ten)
        norm_short = _norm_vn(re.sub(r'\(.*?\)', '', ten))
        if norm_full:  reverse[norm_full]  = sql
        if norm_short: reverse[norm_short] = sql
    return reverse

def match_col_to_sql(excel_col_norm, reverse_map):
    """Khớp cột Excel (đã norm) → SQL column. None nếu không tìm thấy."""
    if excel_col_norm in reverse_map:
        return reverse_map[excel_col_norm]
    for key, sql_col in reverse_map.items():
        if key and excel_col_norm and (
            excel_col_norm.startswith(key[:6]) or key.startswith(excel_col_norm[:6])
        ):
            return sql_col
    return None

# ─────────────────────────────────────────────────────────────────────────────
# 3. PARSER — cải tiến từ main_v1.py
# ─────────────────────────────────────────────────────────────────────────────

# META_KEYS — fallback khi chưa load mapping.
# Khi mapping đã load, dùng build_meta_keys_from_mapping() thay thế.
META_KEYS = {
    "Dự án":        r"Dự án",
    "Đơn hàng":     r"Đơn hàng",
    "Tên sản phẩm": r"Tên [Ss]ản phẩm|Sản phẩm",
    "Mã sản phẩm":  r"Mã sản phẩm|Mã SP",
    "Số lượng":     r"Số lượng",
    "Mục số":       r"Mục số",
    "Hoàn thiện":   r"Hoàn thiện",
    "Kích thước":   r"Kích thước",
    "Kết cấu":      r"Kết cấu",
    "Vật liệu":     r"Vật liệu",
}

def build_meta_keys_from_mapping(mapping):
    """
    Build META_KEYS động từ HEADER mapping.
    Mỗi trường Nguon_DL=Excel có Ten_Excel → dùng làm key + regex pattern.
    Fallback sang META_KEYS cứng nếu mapping trống.
    """
    keys = {}
    for rec in mapping.get('HEADER', []):
        if rec.get('nguon_dl') != 'Excel':
            continue
        ten = rec.get('ten_excel', '').strip()
        if not ten:
            continue
        if '|' in ten:
            # Multi-field: đăng ký từng phần riêng (bỏ @FieldName vì đó là ref parent_row)
            for part in ten.split('|'):
                part = part.strip()
                if part and not part.startswith('@'):
                    keys[part] = re.escape(part)
        else:
            keys[ten] = re.escape(ten)
    return keys if keys else META_KEYS

HEADER_ANCHORS    = ["MÃ SP", "STT", "Tên chi tiết", "Tên Vật Tư", "Mã chi tiết", "Mã vật tư"]
SECTION_STT_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9]*(\.[0-9]+)?\.?$")  # A, B, E1, E2, I., II., III., A.1...
NUMERIC_STT_PATTERN = re.compile(r"^\d+(\.\d+)*$")               # 1, 2, 1.1, 2.3...
_ROMAN_SIMPLE_RE = re.compile(
    r'^(X{0,3})(IX|IV|V?I{0,3})\.?$', re.IGNORECASE)


