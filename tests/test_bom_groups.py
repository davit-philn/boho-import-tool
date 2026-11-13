import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'Tools'))

from services.bom_groups import (
    build_group_table, resolve_group, is_group_stt, is_roman_stt, is_sub_group_row)
from services.mapping_loader import load_mapping
from services.utils import _norm_vn
from services.validators import validate_layer1


def _table():
    return build_group_table(load_mapping(), _norm_vn)


def test_group_table_loaded_from_mapping():
    tb = _table()
    assert {'BOM2', 'BOM3', 'BOM4', 'BOM5'} <= set(tb)
    assert len(tb['BOM3']) >= 7


def test_resolve_by_name_for_both_templates():
    tb = _table()
    # mẫu mới (H) và mẫu cũ (E) cùng quy về loại H
    new = resolve_group('BOM3', [None, None, 'H', 'THẦU PHỤ GIA CÔNG'], tb, _norm_vn)
    old = resolve_group('BOM3', [None, None, 'E', 'Thầu phụ gia công'], tb, _norm_vn)
    assert new['loai_chuan'] == old['loai_chuan'] == 'H' and new['cach_xu_ly'] == 'SP'


def test_resolve_old_template_letters_not_confused():
    tb = _table()
    assert resolve_group('BOM3', ['A', 'NHÓM CHỈ CẠNH'], tb, _norm_vn)['loai_chuan'] == 'D'
    assert resolve_group('BOM3', ['B.2', 'GIA CÔNG KIM LOẠI'], tb, _norm_vn)['loai_chuan'] == 'E2'
    assert resolve_group('BOM3', ['B.1', 'VẬT TƯ KIM LOẠI'], tb, _norm_vn)['ma_tam'] == 'MKT_COKHI'
    assert resolve_group('BOM3', ['C', 'VẬT TƯ PHỤ'], tb, _norm_vn)['ma_tam'] == 'MKT_VTP'


def test_resolve_unknown_and_section_scoped():
    tb = _table()
    assert resolve_group('BOM3', ['Z', 'KÍNH CƯỜNG LỰC'], tb, _norm_vn) is None
    assert resolve_group('BOM2', ['D', 'NHÓM CHỈ CẠNH'], tb, _norm_vn) is None   # nhóm này chỉ khai báo cho BOM3
    assert resolve_group('BOM4', ['STT Kiện', 'F', 'BAO BÌ'], tb, _norm_vn)['loai_chuan'] == 'I'


def test_stt_helpers():
    assert is_group_stt('A') and is_group_stt('E1') and is_group_stt('B.1')
    assert not is_group_stt('1.1') and not is_group_stt(None) and not is_group_stt('nan')
    assert is_roman_stt('II') and is_roman_stt('III') and not is_roman_stt('A') and not is_roman_stt('')


def test_sub_group_row_has_dimensions():
    assert is_sub_group_row([('STT', 'I'), ('DÀY', 50.0)], _norm_vn)
    assert not is_sub_group_row([('STT', 'A'), ('DÀY', None), ('RỘNG', 0)], _norm_vn)


def _errors(rows, section='BOM3', cols=('STT', 'Tên vật tư', 'ĐVT')):
    df = pd.DataFrame(rows, columns=list(cols))
    tables = {'BOM Phần III': {'df': df, 'type': section, 'warnings': []}}
    res = validate_layer1(tables, {'Mục số': '1'}, load_mapping())
    return [e['message'] for e in res['BOM Phần III'] if 'chưa khai báo' in e['message']]


def test_validator_flags_unknown_group_only():
    rows = [['D', 'NHÓM CHỈ CẠNH', None], ['1', 'Chỉ nhựa', 'M'],
            ['Z', 'KÍNH CƯỜNG LỰC', None], ['1', 'Kính', 'TẤM']]
    errs = _errors(rows)
    assert len(errs) == 1 and 'KÍNH CƯỜNG LỰC' in errs[0]


def test_validator_ignores_roman_sub_group():
    rows = [['D', 'NHÓM CHỈ CẠNH', None], ['II', 'TỰA LƯNG', None], ['1', 'Chỉ', 'M']]
    assert _errors(rows) == []
