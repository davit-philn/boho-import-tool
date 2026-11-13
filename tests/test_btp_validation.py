import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'Tools'))

from services.mapping_loader import load_mapping
from services.validators import validate_layer1

COLS = ['STT', 'Tên chi tiết', 'Tên vật tư', 'ĐVT']


def _errors(rows, meta=None):
    df = pd.DataFrame(rows, columns=COLS)
    tables = {'BOM Phần II': {'df': df, 'type': 'BOM2', 'warnings': []}}
    res = validate_layer1(tables, meta if meta is not None else {'Mục số': '11'}, load_mapping())
    return [e['message'] for e in res['BOM Phần II'] if e['severity'] == 'error']


def test_btp_valid_has_no_error():
    rows = [['A', 'VẬT TƯ GỖ, VẬT TƯ TẤM', None, None],
            ['1', 'KHUNG BAO', None, 'CỤM'],
            ['2', 'Panel', 'MDF + Foam', 'cụm']]
    assert _errors(rows) == []


def test_btp_missing_ten_chi_tiet():
    rows = [['A', 'VẬT TƯ GỖ, VẬT TƯ TẤM', None, None], ['1', None, None, 'CỤM']]
    errs = _errors(rows)
    assert len(errs) == 1 and 'Tên chi tiết' in errs[0]


def test_btp_without_group_row():
    errs = _errors([['1', 'KHUNG BAO', None, 'CỤM']])
    assert len(errs) == 1 and 'nhóm nào' in errs[0]


def test_btp_missing_header_muc_so():
    rows = [['A', 'VẬT TƯ GỖ, VẬT TƯ TẤM', None, None], ['1', 'KHUNG BAO', None, 'CỤM']]
    errs = _errors(rows, meta={'Mục số': ''})
    assert len(errs) == 1 and 'Mục số' in errs[0]


def test_non_cum_rows_only_get_required_unit_error():
    rows = [['A', 'VẬT TƯ GỖ, VẬT TƯ TẤM', None, None],
            ['1', 'KHUNG BAO', None, None],            # Đvt trống → lỗi bắt buộc
            ['2', 'Panel', 'MDF + Foam', 'TẤM']]       # Đvt khác CỤM → không báo gì
    errs = _errors(rows)
    assert len(errs) == 1 and 'để trống' in errs[0]


def test_cum_row_with_regular_name_is_btp_with_warning_not_error():
    rows = [['A', 'VẬT TƯ GỖ, VẬT TƯ TẤM', None, None],
            ['1', 'KHUNG BAO', 'MDF E1', 'CỤM']]
    df = pd.DataFrame(rows, columns=COLS)
    tables = {'BOM Phần II': {'df': df, 'type': 'BOM2', 'warnings': []}}
    res = validate_layer1(tables, {'Mục số': '11'}, load_mapping())['BOM Phần II']
    assert [e for e in res if e['severity'] == 'error'] == []
    warns = [e['message'] for e in res if e['severity'] == 'warning']
    assert len(warns) == 1 and 'CỤM' in warns[0]
