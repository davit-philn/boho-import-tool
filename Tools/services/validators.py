"""
services/validators.py — Validate layer 1: kiểu dữ liệu, required fields, length.
"""
import re
import datetime
import pandas as pd

from services.utils import _norm_vn
from services.mapping_loader import (
    build_reverse_map, match_col_to_sql, SECTION_STT_PATTERN, NUMERIC_STT_PATTERN)
from services.bom_rules import find_col, btp_row_problem, btp_row_warning, row_is_btp
from services.bom_groups import (
    build_group_table, resolve_group, is_group_stt, is_sub_group_row, is_roman_stt)
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
                    "row": "Đầu phiếu", "col": ten,
                    "message": f"{ten}: để trống",
                    "severity": "error"
                })
    else:
        # fallback khi chưa load mapping
        for field in REQUIRED_HEADER_FIELDS:
            val = global_meta.get(field, '')
            if not str(val).strip():
                h_errors.append({
                    "row": "Đầu phiếu", "col": field,
                    "message": f"{field}: để trống",
                    "severity": "error"
                })
    result["[H] Đầu phiếu"] = h_errors

    # ── Detail tables ─────────────────────────────────────────────────────────
    group_table = build_group_table(mapping, _norm_vn)   # {} nếu chưa khai báo NHOM_QUY_DOI
    for label, tbl in tables.items():
        if label == "[H] Đầu phiếu":
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
                        "message": f"Thiếu cột {rec['ten_excel']}",
                        "severity": "error"
                    })
                    continue
                # Bảng BOM: dòng tiêu đề nhóm (STT = A, B, II…) không có dữ liệu vật tư → không bắt buộc
                stt_col = (next((c for c in df.columns if _norm_vn(str(c)) == "stt"), None)
                           if section.startswith("BOM") else None)
                blank_rows = []
                for idx, val in enumerate(df[matched_col]):
                    if stt_col is not None and is_group_stt(df[stt_col].iloc[idx]):
                        continue
                    v = "" if (val is None or
                               (isinstance(val, float) and pd.isna(val))) \
                        else str(val).strip()
                    if not v:
                        blank_rows.append(idx)
                n_data = sum(1 for i in range(len(df))
                             if stt_col is None or not is_group_stt(df[stt_col].iloc[i]))
                if len(blank_rows) >= 2 and len(blank_rows) == n_data:
                    # Cả cột trống → thường là cột công thức trong file chưa được Excel tính lại
                    errors.append({
                        "row": "—", "col": matched_col,
                        "message": (f"Cột {str(matched_col).replace('_', ' ')} trống hoàn toàn — "
                                    f"mở file bằng Excel, lưu lại rồi nạp lại"),
                        "severity": "error"
                    })
                for idx in blank_rows:
                    errors.append({
                        "row": idx + 1, "col": matched_col,
                        "message": f"{str(matched_col).replace('_', ' ')}: để trống",
                        "severity": "error"
                    })

        # ── Check 2: Kiểu số ──────────────────────────────────────────────
        # Loại trừ cột có chứa "Tên"/"Name" (text) hoặc "Lọc"/"Loc" (flag IN/OUT)
        _NUMERIC_EXCLUDE_RE = re.compile(
            r"(Tên|Name|Lọc|Loc|_IN\b|_OUT\b)", re.IGNORECASE)
        # Giá trị placeholder (dấu gạch ngang) → coi như NULL, bỏ qua
        _PLACEHOLDER_RE = re.compile(r"^-+$")
        for col in df.columns:
            if mapping and col not in col_info:
                continue   # có mapping nhưng cột chưa mapped → bỏ qua
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
                        "message": f"{str(col).replace('_', ' ')}: '{val}' không phải số",
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
                        "message": f"{str(col).replace('_', ' ')}: '{val_str}' sai định dạng ngày (dd/mm/yyyy)",
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
                        "message": f"{str(col).replace('_', ' ')}: {len(v_str)} ký tự (tối đa {max_len})",
                        "severity": "warning"
                    })

        # ── Check 5: Dòng BTP của BOM2 (Tên vật tư trống/'+' VÀ Đvt CỤM → cần Tên chi tiết) ──
        if section == "BOM2":
            vt_col  = find_col(df.columns, _norm_vn, "Tên vật tư")
            dvt_col = find_col(df.columns, _norm_vn, "Đvt")
            ct_col  = find_col(df.columns, _norm_vn, "Tên chi tiết")
            stt_c   = find_col(df.columns, _norm_vn, "STT")
            if stt_c is not None:
                has_group, n_btp = False, 0   # has_group: đã gặp dòng nhóm (A, B, C…) phía trên
                for idx in range(len(df)):
                    stt_v = str(df[stt_c].iloc[idx]).strip()
                    if is_group_stt(stt_v):
                        has_group = True
                        continue
                    if not NUMERIC_STT_PATTERN.match(stt_v):
                        continue
                    dvt_v = df[dvt_col].iloc[idx] if dvt_col is not None else None
                    ct_v  = df[ct_col].iloc[idx] if ct_col is not None else None
                    vt_v  = df[vt_col].iloc[idx] if vt_col is not None else None
                    problem = btp_row_problem(section, vt_v, dvt_v, ct_v, _norm_vn)
                    warn = btp_row_warning(section, vt_v, dvt_v, _norm_vn)
                    if warn:
                        errors.append({
                            "row": idx + 1, "col": dvt_col or vt_col,
                            "message": warn,
                            "severity": "warning"
                        })
                    if not problem and row_is_btp(section, vt_v, dvt_v, _norm_vn):
                        n_btp += 1
                        if not has_group:   # ItemType lấy từ STT dòng nhóm — không có thì SP bỏ qua dòng
                            problem = "BTP chưa thuộc nhóm nào"
                    if problem:
                        errors.append({
                            "row": idx + 1, "col": ct_col if "Tên chi tiết" in problem else (dvt_col or vt_col),
                            "message": problem,
                            "severity": "error"
                        })
                # Mục số (Header) dùng để đặt/tái dùng mã BTP dạng ItemCode0.MụcSố.000N
                muc_so = {_norm_vn(k): v for k, v in global_meta.items()}.get(_norm_vn("Mục số"))
                if n_btp and not str(muc_so or "").strip():
                    errors.append({
                        "row": "—", "col": "Mục số",
                        "message": "Đầu phiếu thiếu Mục số (cần cho BTP)",
                        "severity": "error"
                    })

        # ── Check 6: Nhóm vật tư phải có trong bảng quy đổi (sheet NHOM_QUY_DOI) ──
        # Quy đổi theo TÊN nhóm (không theo chữ) để mẫu cũ/mới đều nhận ra; nhóm lạ báo lỗi
        # thay vì bị gán mã tạm sai.
        if section.startswith("BOM") and group_table.get(section):
            stt_g = find_col(df.columns, _norm_vn, "STT")
            if stt_g is not None:
                for idx in range(len(df)):
                    if not is_group_stt(df[stt_g].iloc[idx]):
                        continue
                    row = df.iloc[idx]
                    if is_sub_group_row(row.items(), _norm_vn):
                        continue   # dòng MODULE… có kích thước = dữ liệu, không phải nhóm
                    if is_roman_stt(df[stt_g].iloc[idx]):
                        continue   # I, II… = nhóm con (MODULE, TỰA LƯNG…), thừa hưởng nhóm chữ phía trên
                    if resolve_group(section, row.values, group_table, _norm_vn) is None:
                        shown = " ".join(str(v).strip() for v in row.values
                                         if isinstance(v, str) and v.strip())[:50]
                        errors.append({
                            "row": idx + 1, "col": stt_g,
                            "message": f"Nhóm '{shown}' chưa khai báo",
                            "severity": "error"
                        })

        result[label] = errors

    return result

def count_errors(val_errors):
    n_err = sum(len([e for e in v if e["severity"] == "error"])  for v in val_errors.values())
    n_wrn = sum(len([e for e in v if e["severity"] == "warning"]) for v in val_errors.values())
    return n_err, n_wrn

