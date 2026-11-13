"""Quy tắc BOM thuần (không phụ thuộc UI/DB) — dễ test."""

# Nhóm vật tư mà usp_B20BOM_Create_ItemCode tự tạo mã BTP
_SP_BTP_TYPES = {
    'BOM3': {'H', 'E2'},
    'BOM2': {'H'},
}

# Giá trị placeholder trong ô Tên vật tư — coi như trống
_PLACEHOLDERS = {'_', '--', '-', 'x', 'n/a'}

# Đvt đánh dấu dòng BTP (so sánh sau khi bỏ dấu/hoa thường: CỤM = Cụm = cụm)
_BTP_UNIT_NORM = 'cum'


def sp_expects_code(section, item_type, btp_row=False):
    """True nếu dòng PHẢI nhận mã BTP từ SP (không dùng MKT, không để trống).

    btp_row: dòng khớp rule BTP của BOM2 (xem row_is_btp).
    """
    if btp_row:
        return True
    return str(item_type or '').strip() in _SP_BTP_TYPES.get(section, set())


def find_col(columns, norm_fn, label):
    """Tìm cột theo tên chuẩn hóa: khớp đúng trước, rồi khớp đuôi (cột bị ghép header
    cha, vd 'SLg_Tên_Vật_Tư' cho 'Tên vật tư'). None nếu không có."""
    want = norm_fn(label)
    cols = list(columns)
    for c in cols:
        if norm_fn(str(c)) == want:
            return c
    for c in cols:
        n = norm_fn(str(c))
        if n.endswith(want) and len(n) > len(want):
            return c
    return None


def _text(value):
    """Chuỗi đã strip; None/NaN → ''."""
    if value is None:
        return ''
    s = str(value).strip()
    return '' if s.lower() == 'nan' else s


def _name_is_blank(ten_vat_tu):
    s = _text(ten_vat_tu)
    return not s or s.lower() in _PLACEHOLDERS


def btp_name_candidate(ten_vat_tu):
    """Tên vật tư trống (hoặc placeholder) hoặc có dấu '+' (ghép nhiều vật liệu)."""
    return _name_is_blank(ten_vat_tu) or '+' in _text(ten_vat_tu)


def row_is_btp(section, ten_vat_tu, dvt, norm_fn):
    """Dòng BTP (chỉ BOM2): Đvt = CỤM. Tên vật tư không còn là điều kiện (tham số giữ lại cho
    tương thích chỗ gọi) — người dùng đã chủ động ghi CỤM thì dòng đó là BTP."""
    return section == 'BOM2' and norm_fn(_text(dvt)) == _BTP_UNIT_NORM


def btp_row_warning(section, ten_vat_tu, dvt, norm_fn):
    """Cảnh báo (không chặn) khi dòng CỤM lại có Tên vật tư thường (không trống, không có '+'):
    có thể ghi nhầm Đvt — dòng sẽ tạo mã BTP và bỏ qua Tên vật tư thay vì tra mã vật tư."""
    if row_is_btp(section, ten_vat_tu, dvt, norm_fn) and not btp_name_candidate(ten_vat_tu):
        return "Đvt CỤM nhưng có Tên vật tư — sẽ tạo mã BTP, bỏ qua Tên vật tư"
    return None


def btp_row_problem(section, ten_vat_tu, dvt, ten_chi_tiet, norm_fn):
    """Thông báo lỗi nếu dòng BTP (row_is_btp) không đủ điều kiện tạo mã; None nếu ổn.

    Chỉ xét dòng đã là BTP (Đvt = CỤM). Đvt khác CỤM → không phải BTP, không báo gì ở đây
    (Đvt trống đã có lỗi "trường bắt buộc" riêng).
    """
    if not row_is_btp(section, ten_vat_tu, dvt, norm_fn):
        return None
    if _name_is_blank(ten_chi_tiet):
        return "BTP thiếu Tên chi tiết"
    return None


def row_is_bb(section, unit, norm_fn):
    """Dòng BOM2 có Đvt = BB (bao bì) → IsBB = 1. Không phân biệt hoa/thường/khoảng trắng."""
    return section == 'BOM2' and norm_fn(_text(unit)) == 'bb'


def find_missing_sp_codes(section, rows_by_idx, btp_idx=frozenset()):
    """Trả về [(idx, row_vals)] của các dòng SP phải tạo mã nhưng ItemId vẫn rỗng.

    rows_by_idx: {idx: row_vals} — các dòng đã gửi cho SP và đã map kết quả về.
    btp_idx: chỉ số các dòng khớp rule BTP của BOM2.
    """
    missing = []
    for idx, row_vals in rows_by_idx.items():
        if row_vals.get('ItemId'):
            continue
        if sp_expects_code(section, row_vals.get('ItemType'),
                           btp_row=idx in btp_idx):
            missing.append((idx, row_vals))
    return missing
