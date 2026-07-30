"""
services/validators.py — Validate layer 1: kiểu dữ liệu, required fields, length.
"""
import re, datetime
import pandas as pd

from services.utils import _norm_vn
from services.mapping_loader import build_reverse_map, match_col_to_sql
# ─────────────────────────────────────────────────────────────────────────────
# 4. VALIDATE LAYER 1
# ─────────────────────────────────────────────────────────────────────────────

REQUIRED_HEADER_FIELDS = ["Mã sản phẩm", "Dự án", "Đơn hàng"]

NUMERIC_COL_PATTERNS = [
    r"(Dày|Rộng|Dài|Thickness|Width|Length)",
    r"(Số lượng|Quantity|SLg|SLSP)",
    r"(Khối lượng|Weight|WeightCalc)",
    r"(Chiều dài|EdgeLength)",
    r"(Diện tích|Area|FinishArea|PrimerArea)",
    r"(Số cạnh|EdgeCount)",
    r"(Số mặt|FinishSide|PrimerSide)",
]
NUMERIC_RE = re.compile("|".join(NUMERIC_COL_PATTERNS), re.IGNORECASE)

_DATE_FMTS = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%y"]

def _try_parse_date(s, fmt):
    try:
        datetime.datetime.strptime(s, fmt)
        return True
    except (ValueError, TypeError):
        return False

def validate_layer1(tables, global_meta, mapping=None):
    """
    Validate cấu trúc dữ liệu.
    Khi có mapping, kiểm tra thêm: trường bắt buộc, datetime format, nvarchar length.
    Trả về: { table_label: [ {row, col, message, severity} ] }
    """
    result = {}

    # ── Header — lấy danh sách bắt buộc từ mapping (Bat_buoc=CÓ, Nguon_DL=Excel) ───
    h_errors = []
    norm_meta_h = {_norm_vn(k): v for k, v in global_meta.items()}
    if mapping and mapping.get('HEADER'):
        required_recs = [
            r for r in mapping['HEADER']
            if r.get('bat_buoc', '').rstrip('0').rstrip('.') == '1'
            and r.get('nguon_dl', '') == 'Excel'
        ]
        for rec in required_recs:
            ten = rec.get('ten_excel', '').strip()
            if not ten:
                continue
            val = (global_meta.get(ten)
                   or norm_meta_h.get(_norm_vn(ten))
                   or '')
            if not str(val).strip():
                h_errors.append({
                    "row": "Header", "col": ten,
                    "message": f"Trường bắt buộc '{ten}' đang trống",
                    "severity": "error"
                })
    else:
        # fallback khi chưa load mapping
        for field in REQUIRED_HEADER_FIELDS:
            val = global_meta.get(field, '')
            if not str(val).strip():
                h_errors.append({
                    "row": "Header", "col": field,
                    "message": f"Trường bắt buộc '{field}' đang trống",
                    "severity": "error"
                })
    result["[H] Phiếu Header"] = h_errors

    # ── Detail tables ─────────────────────────────────────────────────────────
    for label, tbl in tables.items():
        if label == "[H] Phiếu Header":
            continue
        df      = tbl["df"]
        section = tbl["type"]   # BOM2 / BOM3 / BOM4
        errors  = []

        if df.empty or "Lỗi" in df.columns:
            result[label] = errors
            continue

        # Xây col_info: {excel_col → record từ mapping}
        col_info = {}
        if mapping:
            rev_map  = build_reverse_map(mapping, section)
            sec_recs = {r["sql_col"]: r for r in mapping.get(section, [])}
            for col in df.columns:
                sql_col = match_col_to_sql(_norm_vn(col), rev_map)
                if sql_col and sql_col in sec_recs:
                    col_info[col] = sec_recs[sql_col]

        # ── Check 1: Trường bắt buộc từ mapping ──────────────────────────
        if mapping:
            for rec in mapping.get(section, []):
                if rec["bat_buoc"].rstrip("0").rstrip(".") != "1" or rec["nguon_dl"] != "Excel":
                    continue
                matched_col = next(
                    (c for c, info in col_info.items()
                     if info["sql_col"] == rec["sql_col"]), None)
                if matched_col is None:
                    errors.append({
                        "row": "—", "col": rec["sql_col"],
                        "message": (f"Thiếu cột bắt buộc: '{rec['ten_excel']}'"
                                    f"  (SQL: {rec['sql_col']})"),
                        "severity": "error"
                    })
                    continue
                for idx, val in enumerate(df[matched_col]):
                    v = "" if (val is None or
                               (isinstance(val, float) and pd.isna(val))) \
                        else str(val).strip()
                    if not v:
                        errors.append({
                            "row": idx + 1, "col": matched_col,
                            "message": (f"Dòng {idx+1} | '{matched_col}'"
                                        f": trường bắt buộc đang trống"),
                            "severity": "error"
                        })

        # ── Check 2: Kiểu số ──────────────────────────────────────────────
        # Loại trừ cột có chứa "Tên"/"Name" (text) hoặc "Lọc"/"Loc" (flag IN/OUT)
        _NUMERIC_EXCLUDE_RE = re.compile(
            r"(Tên|Name|Lọc|Loc|_IN\b|_OUT\b)", re.IGNORECASE)
        # Giá trị placeholder (dấu gạch ngang) → coi như NULL, bỏ qua
        _PLACEHOLDER_RE = re.compile(r"^-+$")
        for col in df.columns:
            if not NUMERIC_RE.search(col):
                continue
            if _NUMERIC_EXCLUDE_RE.search(col):
                continue   # cột flag/text ẩn sau tên có keyword số
            for idx, val in enumerate(df[col]):
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    continue
                val_str = str(val).strip()
                if not val_str or _PLACEHOLDER_RE.match(val_str):
                    continue   # trống hoặc dấu gạch (--, -, ---) → bỏ qua
                try:
                    float(val_str.replace(",", "."))
                except (ValueError, TypeError):
                    errors.append({
                        "row": idx + 1, "col": col,
                        "message": (f"Dòng {idx+1} | '{col}'"
                                    f": '{val}' không phải số"),
                        "severity": "error"
                    })

        # ── Check 3: Kiểu datetime ────────────────────────────────────────
        for col, info in col_info.items():
            if info["kieu_dl"] != "datetime":
                continue
            for idx, val in enumerate(df[col]):
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    continue
                if isinstance(val, (datetime.datetime, datetime.date)):
                    continue   # openpyxl đã parse → OK
                val_str = str(val).strip()
                if not val_str:
                    continue
                if not any(_try_parse_date(val_str, f) for f in _DATE_FMTS):
                    errors.append({
                        "row": idx + 1, "col": col,
                        "message": (f"Dòng {idx+1} | '{col}'"
                                    f": '{val_str}' không đúng định dạng ngày"
                                    f" (dd/mm/yyyy)"),
                        "severity": "warning"
                    })

        # ── Check 4: char/varchar/nvarchar length ─────────────────────────
        for col, info in col_info.items():
            if info["kieu_dl"] not in ("nvarchar", "varchar", "char", "nchar") \
                    or not info.get("do_dai"):
                continue
            try:
                max_len = int(info["do_dai"])
            except (ValueError, TypeError):
                continue
            for idx, val in enumerate(df[col]):
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    continue
                v_str = str(val)
                if len(v_str) > max_len:
                    errors.append({
                        "row": idx + 1, "col": col,
                        "message": (f"Dòng {idx+1} | '{col}'"
                                    f": {len(v_str)} ký tự > max {max_len}"),
                        "severity": "warning"
                    })

        result[label] = errors

    return result

def count_errors(val_errors):
    n_err = sum(len([e for e in v if e["severity"] == "error"])  for v in val_errors.values())
    n_wrn = sum(len([e for e in v if e["severity"] == "warning"]) for v in val_errors.values())
    return n_err, n_wrn

