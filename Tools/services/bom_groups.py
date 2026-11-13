"""Nhóm vật tư BOM (A, B, C… hoặc D…H) — quy đổi theo TÊN NHÓM, không theo chữ.

Mẫu BOM cũ và mới dùng chữ khác nhau cho cùng một nhóm (vd "Thầu phụ gia công" là E
ở mẫu cũ, H ở mẫu mới). Bảng NHOM_QUY_DOI trong CK_Mapping_v5.xlsx khai báo tên nhóm →
loại chuẩn (ItemType) + mã tạm MKT. Module thuần, không phụ thuộc UI/DB.
"""

import re

from services.mapping_loader import SECTION_STT_PATTERN

_ROMAN_RE = re.compile(r'^(?=[IVXivx])(X{0,3})(IX|IV|V?I{0,3})\.?$', re.IGNORECASE)

# Cách xử lý khai báo trong bảng
HOW_MKT, HOW_SP, HOW_PARENT = 'MKT', 'SP', 'NHOM_CHA'


def _blank(v):
    return v is None or str(v).strip() in ('', 'nan', 'None')


def is_group_stt(value):
    """STT dạng chữ (A, B, E1, II…) = dòng nhóm. None/NaN không tính."""
    return (not _blank(value)) and bool(SECTION_STT_PATTERN.match(str(value).strip()))


def is_roman_stt(value):
    """STT số La Mã (I, II, III…) = nhóm con (MODULE, TỰA LƯNG…) nằm trong nhóm chữ phía trên."""
    return (not _blank(value)) and bool(_ROMAN_RE.match(str(value).strip()))


def build_group_table(mapping, norm_fn):
    """{section: [(tên nhóm chuẩn hóa, record)]} — tên dài xếp trước để khớp dài nhất thắng."""
    table = {}
    for rec in (mapping or {}).get('NHOM_QUY_DOI', []):
        key = norm_fn(rec.get('ten_nhom', ''))
        if key:
            table.setdefault(rec.get('section', '').strip(), []).append((key, rec))
    for lst in table.values():
        lst.sort(key=lambda kv: len(kv[0]), reverse=True)
    return table


def resolve_group(section, row_values, table, norm_fn):
    """Record bảng quy đổi của dòng nhóm; None nếu không khớp tên nhóm nào.

    row_values: các giá trị ô của dòng nhóm. Khớp khi tên nhóm (bỏ dấu, hoa/thường) nằm
    trong chữ của dòng — nên 'D  NHÓM CHỈ CẠNH', 'STT Kiện F BAO BÌ' đều khớp.
    """
    text = norm_fn(' '.join(str(v) for v in row_values if isinstance(v, str)))
    for key, rec in table.get(section, []):
        if key in text:
            return rec
    return None


def is_sub_group_row(row_items, norm_fn):
    """Dòng STT chữ nhưng có kích thước vật lý (MODULE MD1…) = dòng dữ liệu, không phải nhóm.

    row_items: iterable (tên cột, giá trị).
    """
    for col, v in row_items:
        if re.search(r'(dai|day|rong|width|length)', norm_fn(str(col))) \
                and not _blank(v) and str(v).strip() not in ('0', '0.0'):
            return True
    return False
