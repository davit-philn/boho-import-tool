import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'Tools'))

from services.bom_rules import (
    sp_expects_code, find_missing_sp_codes, find_col, row_is_btp, btp_row_problem)
from services.utils import _norm_vn


def test_sp_expects_code_by_section_and_type():
    assert sp_expects_code('BOM3', 'H')
    assert sp_expects_code('BOM3', 'E2')
    assert sp_expects_code('BOM2', 'H')
    assert not sp_expects_code('BOM2', 'E2')
    assert not sp_expects_code('BOM3', 'C')
    assert not sp_expects_code('BOM4', 'H')
    assert not sp_expects_code('BOM3', None)


def test_btp_row_overrides_type():
    assert sp_expects_code('BOM2', 'C', btp_row=True)


def test_find_missing_reports_only_expected_rows_without_itemid():
    rows = {
        0: {'ItemType': 'H', 'ItemId': 101},      # đã có mã
        1: {'ItemType': 'H', 'ItemId': None},     # thiếu -> báo
        2: {'ItemType': 'C', 'ItemId': None},     # không thuộc BTP -> bỏ qua
        3: {'ItemType': 'C', 'ItemId': None},     # dòng BTP (rule BOM2) -> báo
    }
    missing = find_missing_sp_codes('BOM3', rows, btp_idx={3})
    assert [i for i, _ in missing] == [1, 3]


def test_find_missing_empty_when_all_coded():
    assert find_missing_sp_codes('BOM3', {0: {'ItemType': 'E2', 'ItemId': 5}}) == []


def test_find_col_exact_then_suffix():
    cols = ['STT', 'SLg_Tên_Vật_Tư', 'ĐVT']
    assert find_col(cols, _norm_vn, 'Tên vật tư') == 'SLg_Tên_Vật_Tư'
    assert find_col(cols, _norm_vn, 'Đvt') == 'ĐVT'
    assert find_col(['Tên vật tư', 'X_Tên_vật_tư'], _norm_vn, 'Tên vật tư') == 'Tên vật tư'
    assert find_col(['STT'], _norm_vn, 'Đvt') is None


def test_row_is_btp_blank_or_plus_with_cum():
    for dvt in ('CỤM', 'Cụm', 'cụm', ' CỤM '):
        assert row_is_btp('BOM2', '', dvt, _norm_vn)
        assert row_is_btp('BOM2', None, dvt, _norm_vn)
        assert row_is_btp('BOM2', '-', dvt, _norm_vn)          # placeholder = trống
        assert row_is_btp('BOM2', 'MDF A + Tấm B', dvt, _norm_vn)


def test_row_is_btp_negative_cases():
    assert row_is_btp('BOM2', 'MDF E1', 'CỤM', _norm_vn)          # có tên thường vẫn là BTP (Đvt quyết định)
    assert not row_is_btp('BOM2', '', 'TẤM', _norm_vn)            # Đvt khác CỤM
    assert not row_is_btp('BOM2', '', '', _norm_vn)
    assert not row_is_btp('BOM3', '', 'CỤM', _norm_vn)            # chỉ BOM2
    assert not row_is_btp('BOM2', float('nan'), None, _norm_vn)


def test_btp_row_problem():
    assert btp_row_problem('BOM2', 'MDF E1', 'TẤM', 'Panel', _norm_vn) is None   # vật tư thường
    assert btp_row_problem('BOM2', '', 'CỤM', 'Khung', _norm_vn) is None         # BTP hợp lệ
    assert btp_row_problem('BOM2', 'A + B', 'CỤM', 'Khung', _norm_vn) is None
    # Đvt khác CỤM → không phải BTP → không có cảnh báo BTP (Đvt trống có lỗi bắt buộc riêng)
    assert btp_row_problem('BOM2', '', '', 'Khung', _norm_vn) is None
    assert btp_row_problem('BOM2', 'A + B', 'TẤM', 'Khung', _norm_vn) is None
    assert 'Tên chi tiết' in btp_row_problem('BOM2', '', 'CỤM', '', _norm_vn)     # thiếu tên mã
    assert 'Tên chi tiết' in btp_row_problem('BOM2', 'A + B', 'cụm', '-', _norm_vn)
    assert btp_row_problem('BOM3', '', 'CỤM', '', _norm_vn) is None               # chỉ BOM2


def test_row_is_bb_only_bom2_unit_bb():
    from services.bom_rules import row_is_bb
    for u in ('BB', 'bb', ' Bb '):
        assert row_is_bb('BOM2', u, _norm_vn)
    assert not row_is_bb('BOM2', 'TẤM', _norm_vn)
    assert not row_is_bb('BOM2', '', _norm_vn)
    assert not row_is_bb('BOM2', None, _norm_vn)
    assert not row_is_bb('BOM3', 'BB', _norm_vn)      # chỉ BOM2


def test_mapping_declares_isbb_default_zero_for_all_detail_sections():
    from services.mapping_loader import load_mapping
    m = load_mapping()
    for sec in ('BOM2', 'BOM3', 'BOM4', 'BOM5'):
        rec = [r for r in m[sec] if r['sql_col'] == 'IsBB']
        assert len(rec) == 1 and rec[0]['nguon_dl'] == 'CoDinh' and rec[0]['mac_dinh'] == '0'


def test_btp_row_warning_only_for_cum_with_regular_name():
    from services.bom_rules import btp_row_warning
    assert 'Tên vật tư' in btp_row_warning('BOM2', 'MDF E1', 'CỤM', _norm_vn)
    assert btp_row_warning('BOM2', '', 'CỤM', _norm_vn) is None            # tên trống → bình thường
    assert btp_row_warning('BOM2', 'A + B', 'CỤM', _norm_vn) is None       # ghép "+" → bình thường
    assert btp_row_warning('BOM2', 'MDF E1', 'TẤM', _norm_vn) is None      # không phải CỤM
    assert btp_row_warning('BOM3', 'MDF E1', 'CỤM', _norm_vn) is None      # chỉ BOM2
